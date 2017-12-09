from sys import argv
from lxml import etree, objectify
from lxml.builder import E
from lxml.builder import ElementMaker
import xml.etree.ElementTree as ET
import shutil
import argparse
import os.path
import mammoth
import zipfile
import inspect
import copy
import html
import re
from copy import deepcopy

# function to check if the input file is valid
def is_valid_file(parser, arg):
  if not os.path.exists(arg):
    parser.error("The file %s does not exist!" % arg)
  else:
    return open(arg, 'r')  # return an open file handle

# defining the program options
parser = argparse.ArgumentParser(description='While using the Mammoth docx converter, add options to preserve source class names and formatting information as attributes.')
parser.add_argument("-i", dest="filename", required=True,
                    help="The name of the file to read.", metavar="FILE",
                    type=lambda x: is_valid_file(parser, x))
parser.add_argument('--map', dest='mapStyles', action='store_true', default=True,
                   help='Create a custom map to preserve the source docx style names as classes in the output HTML. Default is True.')
parser.add_argument('--verbose', dest='preserveFormatting', action='store_true', default=False,
                   help='Preserve any formatting applied to the docx styles as attributes in the output HTML. Default is False.')

args = parser.parse_args()

docxfile = args.filename
fileName = docxfile.name
fileNameNoExt = fileName.split(".")[0]

# an empty dict for our ultimate parsed data
verboseAttrs = {}

# set our modified class name suffix
suffix = 'HEDmod'

# function to read the styles.xml file from within the docx;
# this is used for extracting style names and formatting information.

def getNameAndPath(myfile):
  fileName = myfile.name
  filePath = os.path.splitext(myfile.name)[0]
  return fileName, filePath

def makeZip(myfile):
  fileName, filePath = getNameAndPath(myfile)
  newName = filePath + ".zip"
  shutil.copyfile(fileName, newName)
  return newName

def unZip(myfile):
  fileName, filePath = getNameAndPath(myfile)
  document = zipfile.ZipFile(myfile, 'a')
  document.extractall(filePath)
  document.close()
  return

def zipDocx(path, myfile):
  os.chdir(path)
  zf = zipfile.ZipFile(myfile, "w")
  for dirname, subdirs, files in os.walk("."):
    zf.write(dirname)
    for filename in files:
      zf.write(os.path.join(dirname, filename))
  zf.close()
  return

def getWordStyles(myzip):
  zip = zipfile.ZipFile(myzip)
  xml_content = zip.read('word/styles.xml')
  return xml_content

def getWordText(myzip):
  zip = zipfile.ZipFile(myzip)
  xml_content = zip.read('word/document.xml')
  return xml_content

# get all the formatting attributes for a style
def getAttrs(element, inputKey="data", inputVal="", inputDict={}):
  attrKey = element.tag
  attrKey = inputKey + "-" + attrKey.split("}").pop()
  attributes = element.attrib
  children = list(element)
  attrVal = ""
  if len(attributes) == 0 and len(children) == 0:
    attrVal = "true"
  elif len(attributes) == 1 and element.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"):
    attrVal = element.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
  elif len(attributes) > 0:
    for name, value in sorted(element.items()):
      name = name.split("}").pop()
      attrVal = attrVal + "%s:%r;" % (name, value)
  # loop through sub-children and add as attr
  walkChildren(element, attrKey, "", inputDict)

  return attrKey, attrVal

# walk through the element tree for a style
def walkChildren(element, inputKey="data", inputVal="", inputDict={}):
  children = list(element)
  attrKey = ""
  attrVal = ""
  for child in element:
    attrKey, attrVal = getAttrs(child, inputKey, "", inputDict)
    inputDict[attrKey] = attrVal
  return attrKey, attrVal, inputDict

# bringing together getWordStyles, getAttrs, and walkChildren
# to create the final verboseDict of all style names and their
# formatting information.
def getAllStyles(myfile):
  # parse the incoming XML
  source = getWordStyles(myfile)
  root = etree.fromstring(source)

  allStyles = {}
  # get all paragraph styles
  for style in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type='paragraph']") :
    styleID = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")
    styleID = styleID.replace("(","").replace(")","")
    # reset the dictionary for this style's children
    allAttr = {}
    # get all child elements of the style
    children = list(style)
    # walk the child tree to collect all elements and their attributes into the dictionary
    attrKey, attrVal, allAttr = walkChildren(style, "data", "", allAttr)
    allAttr['data-w-type'] = 'p'
    # add the style ID to the master dictionary with value = the collected child elements
    allStyles[styleID] = allAttr
        
  # get all character styles
  for style in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type='character']") :
    styleID = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")
    # reset the dictionary for this style's children
    allAttr = {}
    # get all child elements of the style
    children = list(style)
    # walk the child tree to collect all elements and their attributes into the dictionary
    attrKey, attrVal, allAttr = walkChildren(style, "data", "", allAttr)
    allAttr['data-w-type'] = 'r'
    # add the style ID to the master dictionary with value = the collected child elements
    allStyles[styleID] = allAttr

  return allStyles

# run this function before style map and getting style defs
def getDirectFormatting(myfile):
  myzip = makeZip(myfile)
  source = getWordText(myzip)
  root = etree.fromstring(source)

  styles_source = getWordStyles(myzip)
  styles_root = etree.fromstring(styles_source)

  # namespace declarations for the element method we'll use later
  WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  w = "{%s}" % WORD_NAMESPACE

  WORD14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"
  w14 = "{%s}" % WORD_NAMESPACE

  NSMAP = {None : WORD_NAMESPACE}

  E = ElementMaker(namespace="http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                   nsmap={'mc' : "http://schemas.openxmlformats.org/markup-compatibility/2006",
                          'r' : "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                          'w' : "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                          'w14' : "http://schemas.microsoft.com/office/word/2010/wordml"})

  newstyles = []
  modcounter = 1

  #create our paraid style
  newstylename = "HED-dataID"
  STYLEOBJ = E.style
  STYLENAMEOBJ = E.name
  RPROBJ = E.rPr

  newstyle = STYLEOBJ(
    STYLENAMEOBJ(),
    RPROBJ()
  )

  newstyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "character")
  newstyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId", newstylename)
  newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name").set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", newstylename)
  styles_root.append(newstyle)

  for para in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
    # get the paragraph id (for mapping back)
    para_id = para.get("{http://schemas.microsoft.com/office/word/2010/wordml}paraId")
    # get all formatting on the P (inside pPr)
    para_format = para.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    formatting = para.xpath(".//w:pPr/w:*[not(self::w:pStyle)]", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
    style = para.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
    if formatting:
      if style is not None:
        # if there are any non-pstyle children of para_format, 
        # then proceed with modifications
        stylename = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
        newstylename = stylename + suffix + str(modcounter)
        currstyle = styles_root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId='" + stylename + "']")
        newstyle = deepcopy(currstyle)
        if newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}basedOn") is not None:
          newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}basedOn").set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", stylename)
        else:
          newbasedon = etree.Element(w + "basedOn", nsmap=NSMAP)
          newbasedon.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", stylename)
          newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name").addnext(newbasedon)
        if newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr") is None:
          newppr = etree.Element(w + "pPr", nsmap=NSMAP)
          newstyle.append(newppr)
      else:
        # add the pStyle element to the para
        if para_format is None:
          newppr = etree.Element(w + "pPr", nsmap=NSMAP)
          para.insert(0, newppr)

        newpstyle = etree.Element(w + "pStyle", nsmap=NSMAP)
        para.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr").append(newpstyle)

        newstylename = suffix + str(modcounter)
        STYLEOBJ = E.style
        STYLENAMEOBJ = E.name
        PPROBJ = E.pPr

        newstyle = STYLEOBJ(
          STYLENAMEOBJ(),
          PPROBJ()
        )

        newstyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "paragraph")

      # set the para stylename to the new stylename
      stylename = para.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
      stylename.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", newstylename)
      # create new style
      newstyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId", newstylename)
      newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name").set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", newstylename)
      for format in formatting:
        if newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr/" + format.tag) is not None:
          currel = newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr/" + format.tag)
          # copy over just the parts of the element that are different from the existing version
          allchildren = format.xpath("w:*", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
          for att in format.attrib:
            currel.set(att, format.attrib[att])
          for child in allchildren:
            if currel.find(child.tag) is not None:
              currchild = currel.find(child.tag)
              for att in child.attrib:
                currchild.set(att, child.attrib[att])
            else:
              currel.append(child)
        elif format.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr":
          currel = newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
          for att in format.attrib:
            currel.set(att, node.attrib[att])
        else:
          newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr").append(format)
      # add new style to list
      stylelist = styles_root.append(newstyle)
      modcounter += 1
    
    for run in para.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"):
      # get all formatting on the P (inside pPr)
      run_format = run.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
      formatting = run.xpath(".//w:rPr/w:*[not(self::w:rStyle)]", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
      style = run.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rStyle")
      if formatting:
        if style is not None:
          # if there are any non-pstyle children of run_format, 
          # then proceed with modifications
          stylename = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
          newstylename = stylename + suffix + str(modcounter)
          currstyle = styles_root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId='" + stylename + "']")
          newstyle = deepcopy(currstyle)
          if newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}basedOn") is not None:
            newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}basedOn").set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", stylename)
          else:
            newbasedon = etree.Element(w + "basedOn", nsmap=NSMAP)
            newbasedon.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", stylename)
            newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name").addnext(newbasedon)
          if newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr") is None:
            newrpr = etree.Element(w + "rPr", nsmap=NSMAP)
            newstyle.append(newrpr)
        else:
          # add the rStyle element to the run
          if run_format is None:
            newrpr = etree.Element(w + "rPr", nsmap=NSMAP)
            run.insert(0, newrpr)

          newrstyle = etree.Element(w + "rStyle", nsmap=NSMAP)
          run.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr").append(newrstyle)

          newstylename = suffix + str(modcounter)
          STYLEOBJ = E.style
          STYLENAMEOBJ = E.name
          RPROBJ = E.rPr

          newstyle = STYLEOBJ(
            STYLENAMEOBJ(),
            RPROBJ()
          )

          newstyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "character")

        # set the run stylename to the new stylename
        stylename = run.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rStyle")
        stylename.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", newstylename)
        # create new style
        newstyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId", newstylename)
        newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name").set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", newstylename)
        for format in formatting:
          if newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr/" + format.tag) is not None:
            currel = newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr/" + format.tag)
            # copy over just the parts of the element that are different from the existing version
            allchildren = format.xpath("w:*", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            for att in format.attrib:
              currel.set(att, format.attrib[att])
            for child in allchildren:
              if currel.find(child.tag) is not None:
                currchild = currel.find(child.tag)
                for att in child.attrib:
                  currchild.set(att, child.attrib[att])
              else:
                currel.append(child)
          elif format.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr":
            currel = newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
            for att in format.attrib:
              currel.set(att, node.attrib[att])
          else:
            newstyle.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr").append(format)
        # add new style to list
        stylelist = styles_root.append(newstyle)
        modcounter += 1
    # add the para id onto the new stylename
    if para.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle") is not None:
      newrun = etree.Element(w + "r", nsmap=NSMAP)
      newrpr = etree.Element(w + "rPr", nsmap=NSMAP)
      newtxt = etree.Element(w + "t", nsmap=NSMAP)
      newrstyle = etree.Element(w + "rStyle", nsmap=NSMAP)

      newrstyle.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "HED-dataID")
      newtxt.text = para_id
      newrpr.append(newrstyle)
      newrun.append(newrpr)
      newrun.append(newtxt)
      para.append(newrun)
             
  os.remove(myzip)
  return root, styles_root

def addID(root):
  for run in root.findall('.//span[@class="HED-dataID"]'):
    myid = run.text
    myparent = run.getparent()
    myparent.set("data-source-id", myid)
    myparent.remove(run)
  return root

# delete the mod suffix from class names
def deleteSuffix(root):
  for el in root.xpath('//*[re:test(@class, "' + suffix + '[0-9]+$")]', namespaces={'re': "http://exslt.org/regular-expressions"}):
    newclass = re.sub(r'' + suffix + '[0-9]+$', '', el.get("class"))
    el.set("class", newclass)

# add the formatting info back to the HTML as attributes on each element
def addAttrs(html, myDict):
  root = etree.HTML(html)
  root = addID(root)
  for style, vals in myDict.items():
    for para in root.findall(".//p[@class='" + style + "']"):
      for key, val in vals.items():
        para.attrib[key] = val
    for run in root.findall(".//span[@class='" + style + "']"):
      for key, val in vals.items():
        run.attrib[key] = val
  deleteSuffix(root)
  newHTML = etree.tostring(root, standalone=True, xml_declaration=True)
  return newHTML

def sanitizeHTML(html):
  root = etree.HTML(html)
  root = addID(root)
  deleteSuffix(root)
  newHTML = etree.tostring(root, standalone=True, xml_declaration=True)
  return newHTML

fobj = open(fileName,'rb')

documentxml, stylesxml = getDirectFormatting(fobj)

unZip(fobj)

fobj.close()

filePath = os.path.splitext(fobj.name)[0]

docfilePath = os.path.join(filePath, "word", "document.xml")
stylesfilePath = os.path.join(filePath, "word", "styles.xml")

# write to a new document
docfile = open(docfilePath, 'wb')
docfile.write(etree.tostring(documentxml, encoding="UTF-8", standalone=True, xml_declaration=True))
docfile.close()

# write to a new document
stylesfile = open(stylesfilePath, 'wb')
stylesfile.write(etree.tostring(stylesxml, encoding="UTF-8", standalone=True, xml_declaration=True))
stylesfile.close()

newZipName = fileName + ".zip"

zipDocx(filePath, newZipName)

fobj = open(newZipName,'rb')

#verboseAttrs = getAllStyles(docxfile)
verboseAttrs = getAllStyles(fobj)

# create the style map if requested
if args.mapStyles == True:
  style_map = '"""'
  for style, vals in verboseAttrs.items():
    sourceName = vals['data-name']
    destName = style
    # mapping paragraphs
    if vals['data-w-type'] == 'p':
      thisMap = "p[style-name='" + sourceName + "'] => p." + destName + ":fresh"
    # mapping runs
    else:
      thisMap = "r[style-name='" + sourceName + "'] => span." + destName
    # write this map to the map file
    style_map = style_map + "\n" + thisMap

  style_map = style_map + '\n"""'

# convert with mammoth
result = mammoth.convert_to_html(fobj, style_map=style_map)
html = result.value
messages = result.messages

fobj.close()

# add the verbose attributes to the output HTML if requested
if args.preserveFormatting == True:
  html = addAttrs(html, verboseAttrs)
else:
  html = sanitizeHTML(html)

outputName = fileNameNoExt + ".html"
outputPath = os.path.join(filePath, outputName)

# write to a new HTML document
output = open(outputPath, 'wb')
output.write(html)
output.close()

# cleanup
os.remove(newZipName)
shutil.rmtree(filePath)
