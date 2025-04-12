import win32com.client as win32
import win32com.client
import os
from docx import Document #useful video about docx https://www.youtube.com/watch?v=sGJBUB8yMIw
from docx.text.run import Run
from docx.text.hyperlink import Hyperlink
import json
from pptx import Presentation
import shutil
from bs4 import BeautifulSoup
import glob
import docx
import docx2txt
import zipfile
import shutil

import re

import OLEFileExtractorLib
                 
import colored

def PrintWarningMessage(text):
    warnig_text = colored.Fore.yellow + text + "\033[0m"
    print(warnig_text)
    
def PrintErrorMessage(text):
    error_text = colored.Fore.red + text + "\033[0m"
    print(error_text)   
    
def PrintSuccessMessage(text):
    error_text = colored.Fore.green + text + "\033[0m"
    print(error_text)    
    
def PrintInformationMessage(text):
    error_text = colored.Fore.black + text + "\033[0m"
    print(error_text)  
    
def PrintMessage(text, cColorASCII):
    error_text = cColorASCII + text + "\033[0m"
    print(error_text)    



def RemoveSignsAndNumbers(strInputString):
        # Define a regular expression pattern to match signs (+ or -) and numbers
        strPattern = r'[+\-\[\]()0-9]'
        
        # Use the re.sub() function to replace all matches with an empty string
        strCleanedString = re.sub(strPattern, '', strInputString)
        
        # Remove spaces from the cleaned string
        strCleanedString = strCleanedString.replace(' ', '')
        
        return strCleanedString

def PathAbsoluteToRelative(strAbsolutePath, strReferencePath):
    """Transforms an absolute path into a relative path relative to a reference path."""
    return os.path.relpath(strAbsolutePath, strReferencePath)

def CopyFileToFolder(strSourceFile, strTargetFolder):
    if not os.path.exists(strTargetFolder):
        os.makedirs(strTargetFolder)
    
    strTargetPath = os.path.join(strTargetFolder, os.path.basename(strSourceFile))
    
    shutil.copy(strSourceFile, strTargetPath)
    print(f"File '{strSourceFile}' copied to '{strTargetPath}'")

def MoveFiles(strSourceFolder, strTargetFolder):
    if not os.path.exists(strTargetFolder):
        os.makedirs(strTargetFolder)
    
    for strFilename in os.listdir(strSourceFolder):
        strSourcePath = os.path.join(strSourceFolder, strFilename)
        strTargetPath = os.path.join(strTargetFolder, strFilename)
        shutil.move(strSourcePath, strTargetPath)
        print(f"Moved '{strFilename}' to '{strTargetPath}'")


def RemoveFolder(strFolderPath):
    try:
        shutil.rmtree(strFolderPath)
        print(f"Folder '{strFolderPath}' and its contents removed.")
    except OSError as e:
        print(f"Error while removing folder: {e}")
        
        
        
def GetFilesInFolder(strFolderPath):
    """
    Get a list of all files in the specified folder.

    Args:
        folder_path (str): The path to the folder.

    Returns:
        List[str]: A list of file names in the folder.
    """
    try:
        # Use os.listdir to get a list of all items (files and directories) in the folder
        arrItems = os.listdir(strFolderPath)

        # Filter out only the files from the list
        arrFiles = [item for item in arrItems if os.path.isfile(os.path.join(strFolderPath, item))]

        return arrFiles
    except OSError as e:
        print(f"Error reading folder '{strFolderPath}': {e}")
        return []

class DocxToJsonParser:
    
    def __init__(self):
        
        self.m_strRoot = os.getcwd()
        self.m_strRoot = self.m_strRoot.replace('\\', '/')

        self.m_strOutputDir = self.m_strRoot
        self.m_strWordDocsDir = self.m_strRoot
        # Find all docx files in the provided folder and its subfolders
        self.m_arrDocxFiles = glob.glob(self.m_strWordDocsDir + "/" + '**/*.docx', recursive=True)

        #os.chdir(self.m_strRoot)
        print('Path changed to: ', self.m_strRoot)
     
    # Globals manipulators (should probably be added to each core gneration scripts)
    def SetRootDir(self, strRootDir):
        self.m_strRoot = strRootDir
        self.m_strPathFromRoot = self.m_strCurrentDirectory.removeprefix(self.m_strParentDirectory)
        
    def GetRootDir(self):
        return self.m_strRoot

    def SetOutputFolder(self, strOuptutPath):
        self.m_strOutputDir = strOuptutPath
        self.m_strOutputDir = self.m_strOutputDir.replace('\\', '/')
        
    def GetOutputFolder(self): 
        return self.m_strOutputDir

    def SetDocxPath(self, strDocxDocumentsPath):
        folder_path = strDocxDocumentsPath
        docx_files = glob.glob(folder_path + "/" + '**/*.docx', recursive=True)
        self.SetDocxFilesList(docx_files)

    def GetDocxPath(self):
        return self.m_strWordDocsDir

    def GetDocxFilesList(self):
        return self.m_arrDocxFiles
    
    def SetDocxFilesList(self, arrNewList):
        self.m_arrDocxFiles = arrNewList

    def ExtractDocxImagesToFolder(self, strPath, strOutputPath):
        # Extract the images to img_folder/

        PrintInformationMessage("Extracting Media from: " + strPath + "to:" + strOutputPath)
        # Unzip the .docx file
        with zipfile.ZipFile(strPath, 'r') as zip_ref:
                zip_ref.extractall(os.path.dirname(strPath) + "/" + 'temp')

        # Check if media folder exists in the docs
        # Move all images from it to the output destination
        if os.path.isdir(os.path.dirname(strPath) + "/" + 'temp/word/media/'):
            MoveFiles(os.path.dirname(strPath) + "/" + 'temp/word/media/', strOutputPath)

        RemoveFolder(os.path.dirname(strPath) + "/" + 'temp')
        PrintSuccessMessage("Media extracted successfully from: " + strPath)
        
    def ExtractDocxMp4ToFolder(self, strPath, strOutputPath):
        # Extract the images to img_folder/

        PrintInformationMessage("Extracting Media from: " + strPath + "to:" + strOutputPath)
        # Unzip the .docx file
        with zipfile.ZipFile(strPath, 'r') as zip_ref:
                zip_ref.extractall(os.path.dirname(strPath) + "/" + 'temp')

        # Check if embeddings folder exists in the docs
        # Move all images from it to the output destination
        if os.path.isdir(os.path.dirname(strPath) + "/" + 'temp/word/embeddings/'):
            arrFiles = GetFilesInFolder(os.path.dirname(strPath) + "/" + 'temp/word/embeddings/')
            
            iCount = 0
            for fileEmbeddedFile in arrFiles:
                strExctractedMp4Name = "Video" + "_" + str(iCount)
                OLEFileExtractorLib.ExctractMP4File(os.path.dirname(strPath) + "/" + 'temp/word/embeddings/' + fileEmbeddedFile, strExctractedMp4Name, strOutputPath)
                iCount += 1
            
        # move_files(os.path.dirname(strPath) + "/" + 'temp/word/media/', strOutputPath)

        RemoveFolder(os.path.dirname(strPath) + "/" + 'temp')
        PrintSuccessMessage("Media extracted successfully from: " + strPath)
        
        
    def ParseDocxElementsToArray(self, strPath, strMediaFolderPath):
        
        PrintInformationMessage("Parsing Docx content from: " + strPath)
        bAnyParsingWarnings = False
        
        arrParsedElements = []
        
        # Open you .docx document
        doc = docx.Document(strPath)

        # Save all 'rId:filenames' relationships in an dictionary named rels
        rels = {}
        for r in doc.part.rels.values():
            if isinstance(r._target, docx.parts.image.ImagePart):
                rels[r.rId] = os.path.basename(r._target.partname)

        iParsedEmbeddedMP4 = 0

        # Then process your text
        iParagraphCount = -1
        for paragraph in doc.paragraphs:
            iParagraphCount +=1
            
            arrImagesFound_LOG = []

            
            iImagesCount = 0
            # check if there is an image in this paragraph          
            for run in paragraph.runs:
                for e in run._r:
                    if 'drawing' in e.tag:
                        if  hasattr(e[0], "graphic"):
                            rId2 = e[0].graphic.graphicData.pic.blipFill.blip.embed
                            image_part = doc.part.related_parts[rId2]
                            # soemthing = run._r[0][0].graphic.graphicData.pic.blipFill.blip.embed
                            
                            # Get image name
                            img_name = os.path.basename(image_part.partname) # partname is a path to image within word with image name
                            
                            arrImagesFound_LOG.append(img_name)
                            
                            subdictionary = {}
                            subdictionary["ParagraphNativeIndex"] = iParagraphCount
                            subdictionary["Type"] = "Image"                   
                            subdictionary["Width"] = image_part.image.px_width
                            subdictionary["Height"] = image_part.image.px_height
                            img_path = strMediaFolderPath + img_name
                            imageRelativePath = PathAbsoluteToRelative(img_path, self.m_strRoot)
                            #img_path = finalText.replace("\\", "/") # correct the path to the image
                            subdictionary["Content"] = imageRelativePath
                            #dictionary["Elements"].append(subdictionary)
                            
                            iImagesCount += 1
                            arrParsedElements.append(subdictionary)

                    if 'object' in e.tag:

                            arrFiles = GetFilesInFolder(strMediaFolderPath)

                            # Get only the extracted mp4 file names
                            arrMP4Files = []
                            for strFileName in arrFiles:
                                strFileExtension = strFileName.split(".")[-1]
                                if strFileExtension == "mp4":
                                    arrMP4Files.append(strFileName)
                            
                            fileNameParsedMP4 = arrMP4Files[iParsedEmbeddedMP4]
                            
                            #rId2 = e[0].graphic.graphicData.pic.blipFill.blip.embed
                        # image_part = doc.part.related_parts[rId2]
                            subdictionary = {}
                            subdictionary["ParagraphNativeIndex"] = iParagraphCount
                            subdictionary["Type"] = "EmbeddedObject"                   
                            subdictionary["ObjectType"] = "MP4"
                            object_path = strMediaFolderPath + fileNameParsedMP4
                            objectRelativePath = PathAbsoluteToRelative( object_path, self.m_strRoot)
                            #img_path = finalText.replace("\\", "/") # correct the path to the image
                            subdictionary["Content"] = objectRelativePath
                            
                            iParsedEmbeddedMP4 += 1
                            arrParsedElements.append(subdictionary)
            if iImagesCount > 1:
                bAnyParsingWarnings = True
                strWarningMessage = "Paragraph[" + str(iParagraphCount) + "]" + "contains multiple images! Every image will be parsed as separate img element. There is no img inlining supported"
                
                PrintWarningMessage(strWarningMessage)
                PrintWarningMessage("ImagesFoundInThisParagraph:")
                
                for strImageName in arrImagesFound_LOG:
                    PrintWarningMessage("Image: " + strImageName)

            # If there are any hyperlinks in the paragraph, separately parse hyperlinks and text runs
            # Solution provided in this post: https://github.com/python-openxml/python-docx/issues/1113
            if paragraph.hyperlinks:
                  for item in paragraph.iter_inner_content():
                    if isinstance(item, Run):
                        subdictionary = {}
                        subdictionary["ParagraphNativeIndex"] = iParagraphCount
                        subdictionary["Type"] = "Text"
                        subdictionary["Style"] = str(paragraph.style.name)
            
                        # Find AlignmentRule
                        strAlignment = RemoveSignsAndNumbers(str(paragraph.alignment))
            
                         # "None" means "LEFT"
                        if strAlignment == "None":
                            strAlignment = "LEFT"
            
                        subdictionary["Alignment"] = strAlignment
                        strFinalText = str(item.text)
                
                        if str(subdictionary["Style"]) == "Caption":
                            subdictionary["Alignment"] = "None"
                    
                        strFinalText = strFinalText.replace("\r", "")# remove any excessive \r at the end of the lines
                        strFinalText = strFinalText.replace("\u2019", "'") # find any \u2019 elements and replace them with "'"
                        strFinalText = strFinalText.replace("\u2022\t", "&#x2022") # Replace a bullet with HTML bullet
                        strFinalText = strFinalText.replace("\u2022", "&#x2022") # Replace a bullet with HTML bullet
                        strFinalText = strFinalText.replace("\u2013", "-") #Replace EN DASH character with a simple minus sign
         
                        subdictionary["Content"] = strFinalText
            
                        arrParsedElements.append(subdictionary)
                    #... do the run thing ...
                    elif isinstance(item, Hyperlink):
                        subdictionary = {}
                        subdictionary["ParagraphNativeIndex"] = iParagraphCount
                        subdictionary["Type"] = "Hyperlink"
                        subdictionary["Style"] = str(paragraph.style.name)
            
                        # Find AlignmentRule
                        strAlignment = RemoveSignsAndNumbers(str(paragraph.alignment))
            
                         # "None" means "LEFT"
                        if strAlignment == "None":
                            strAlignment = "LEFT"
            
                        subdictionary["Alignment"] = strAlignment
                        strFinalText = str(item.address)
                
                        if str(subdictionary["Style"]) == "Caption":
                            subdictionary["Alignment"] = "None"
                    
                        strFinalText = strFinalText.replace("\r", "")# remove any excessive \r at the end of the lines
                        strFinalText = strFinalText.replace("\u2019", "'") # find any \u2019 elements and replace them with "'"
                        strFinalText = strFinalText.replace("\u2022\t", "&#x2022") # Replace a bullet with HTML bullet
                        strFinalText = strFinalText.replace("\u2022", "&#x2022") # Replace a bullet with HTML bullet
                        strFinalText = strFinalText.replace("\u2013", "-") #Replace EN DASH character with a simple minus sign
         
                        subdictionary["Content"] = strFinalText
            
                        arrParsedElements.append(subdictionary)
                        
                        
            #... do the hyperlink thing ...
            # Parse text if it's not null
            # Ignore Captions for now
            else:
             if paragraph.text != '' and paragraph.text != " ": #Uncomment this to remove the support for captions: #and paragraph.style.name != "Caption":
            
                if iImagesCount > 0:
                    bAnyParsingWarnings = True
                    strWarningMessage = "Paragraph[" + str(iParagraphCount) + "]" + "contains text and images! Text and image will be parsed as separate element. Images will be parsed before text!!! If it's not desirable, please put text and images into separate paragraphs"
                    PrintWarningMessage(strWarningMessage)
                    PrintWarningMessage("ImagesFoundInThisParagraph:")
                    for strImageName in arrImagesFound_LOG:
                        PrintWarningMessage("Image: " + strImageName)
                
                subdictionary = {}
                subdictionary["ParagraphNativeIndex"] = iParagraphCount
                subdictionary["Type"] = "Text"
                subdictionary["Style"] = str(paragraph.style.name)
            
                # Find AlignmentRule
                strAlignment = RemoveSignsAndNumbers(str(paragraph.alignment))
            
                # "None" means "LEFT"
                if strAlignment == "None":
                    strAlignment = "LEFT"
            
                subdictionary["Alignment"] = strAlignment
                strFinalText = str(paragraph.text)
                
                if str(subdictionary["Style"]) == "Caption":
                    subdictionary["Alignment"] = "None"
                    
                strFinalText = strFinalText.replace("\r", "")# remove any excessive \r at the end of the lines
                strFinalText = strFinalText.replace("\u2019", "'") # find any \u2019 elements and replace them with "'"
                strFinalText = strFinalText.replace("\u2022\t", "&#x2022") # Replace a bullet with HTML bullet
                strFinalText = strFinalText.replace("\u2022", "&#x2022") # Replace a bullet with HTML bullet
                strFinalText = strFinalText.replace("\u2013", "-") #Replace EN DASH character with a simple minus sign
         
                subdictionary["Content"] = strFinalText
            
                arrParsedElements.append(subdictionary)
            
        if bAnyParsingWarnings == True:
            PrintSuccessMessage("Docx parsed successfully but there were some warnings for: " + strPath)
        else:
            PrintSuccessMessage("Docx parsed successfully from: " + strPath)
        
        return arrParsedElements
                    


        
    
        
    def ParseAllProjectPagesDocxToJson(self):
        arrPathToName = []

        # Save the paths to an array
        for strFilepath in self.m_arrDocxFiles:
            dictOutput = {}
            strFilepath = strFilepath.replace('\\', '/')
            dictOutput["DocumentPath"] = strFilepath
            dictOutput["DocumentName"] = os.path.splitext(os.path.basename(strFilepath))[0] #File name without extension
            arrPathToName.append(dictOutput)

        iCount = 0
        for PathAndName in arrPathToName:
        
            # Data to be written
            # We're gonna stoe the content of a single word doc
            # thiss will be written into the json file
            dictOutput = {}

            #dictionary["DocumentTitle"] = str(doc.core_properties.title)
            dictOutput["DocumentTitle"] = "Report"
            dictOutput["DocxPath"] = PathAndName["DocumentPath"]
            dictOutput["Elements"] = []
            
            # Check if the folder exists
            if not os.path.exists(self.m_strOutputDir + "/" + PathAndName["DocumentName"]):
            # If it doesn't exist, create it
                os.makedirs(self.m_strOutputDir + "/" + PathAndName["DocumentName"])
            #print(f"Folder '{full_path}' created.")
            #   else:
        # print(f"Folder '{full_path}' already exists.")
            
            self.ExtractDocxImagesToFolder(PathAndName["DocumentPath"], self.m_strOutputDir + "/" + PathAndName["DocumentName"] + "/" + "Media/")
            self.ExtractDocxMp4ToFolder(PathAndName["DocumentPath"], self.m_strOutputDir + "/" + PathAndName["DocumentName"] + "/" + "Media/")
            dictOutput["Elements"] = (self.ParseDocxElementsToArray(PathAndName["DocumentPath"], self.m_strOutputDir + "/" + PathAndName["DocumentName"] + "/" + "Media/"))
            
            iCount += 0
            json_object = json.dumps(dictOutput, indent=4)
            
            # Writing to sample.json
            with open(self.m_strOutputDir + "/" + PathAndName["DocumentName"] + "/" +  PathAndName["DocumentName"] + ".json", "w") as outfile:
                outfile.write(json_object)
            







