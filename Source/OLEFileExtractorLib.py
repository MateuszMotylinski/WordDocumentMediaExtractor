from zipfile import ZipFile
from glob import glob
from enum import Enum
import sys, os


# Define a file header id in hex enum
class FileHeaderIDHex(Enum):
    MP4 = b'\x00\x00\x00\x20\x66\x74\x79\x70\x69\x73\x6F\x6D'

#Simple lib used for extracting various files from OLE binary files

def ExtractFileWithSignatureToFolder(strOLEFilePath, eFileHeaderIDHex, strOutputFileName, strOutputFileExtension, strOutputFolderPath):
    # Open the binary file in read mode ('rb')
    input_file_path = strOLEFilePath  # Replace with your binary file path
    hex_sequence_to_find =  b'\x00\x00\x00\x20\x66\x74\x79\x70\x69\x73\x6F\x6D'    #eFileHeaderIDHex.value#b'\x00\x00\x00\x20\x66\x74\x79\x70\x69\x73\x6F\x6D'    #b'\x66\x74\x79\x70'     #b'\x00\x00\x00\x00\x20\x66\x74\x79\x70\x69\x73\x6F\x6D'            #b'\x66\x74\x79\x70\x69\x73\x6F\x6D'  # Hexadecimal sequence to find

    try:
        with open(input_file_path, 'rb') as input_file:
            # Read the binary data from the input file
            binary_data = input_file.read()

            # Search for the hexadecimal sequence
            sequence_index = binary_data.find(hex_sequence_to_find)

            if sequence_index != -1:
                # Create a new file and write the found sequence and the rest of the data
                output_file_path = strOutputFolderPath + "/" + strOutputFileName + "." + strOutputFileExtension  # Replace with your desired output file path
                with open(output_file_path, 'wb') as output_file:
                    output_file.write(binary_data[sequence_index:])
                print(f"Hexadecimal sequence found at index {sequence_index} and copied to '{output_file_path}'")
            else:
                print("Hexadecimal sequence not found in the file.")
    except FileNotFoundError:
        print(f"File '{input_file_path}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")



def ExctractMP4File(strOLEFilePath, strOutputFileName, strOutputFolderPath):
    ExtractFileWithSignatureToFolder(strOLEFilePath, FileHeaderIDHex.MP4, strOutputFileName, "mp4", strOutputFolderPath)








""" TEST

#Test
pathname = os.path.dirname(sys.argv[0])        

current_directory = pathname #os.getcwd()
parent_directory = os.path.dirname(current_directory)
g_strRoot =  parent_directory + "/"

g_strPathFromRoot = current_directory.removeprefix(parent_directory)
g_strPathFromRoot = g_strPathFromRoot.replace('\\', '') + "/"

os.chdir(os.getcwd())
print('Path changed to: ', os.getcwd())


# Open the binary file in read mode ('rb')
input_file_path = "oleObject1.bin"  # Replace with your binary file path
hex_sequence_to_find =  b'\x00\x00\x00\x20\x66\x74\x79\x70\x69\x73\x6F\x6D'    #b'\x66\x74\x79\x70'     #b'\x00\x00\x00\x00\x20\x66\x74\x79\x70\x69\x73\x6F\x6D'            #b'\x66\x74\x79\x70\x69\x73\x6F\x6D'  # Hexadecimal sequence to find

try:
    with open(input_file_path, 'rb') as input_file:
        # Read the binary data from the input file
        binary_data = input_file.read()

        # Search for the hexadecimal sequence
        sequence_index = binary_data.find(hex_sequence_to_find)

        if sequence_index != -1:
            # Create a new file and write the found sequence and the rest of the data
            output_file_path = "output_binary_data.bin"  # Replace with your desired output file path
            with open(output_file_path, 'wb') as output_file:
                output_file.write(binary_data[sequence_index:])
            print(f"Hexadecimal sequence found at index {sequence_index} and copied to '{output_file_path}'")
        else:
            print("Hexadecimal sequence not found in the file.")
except FileNotFoundError:
    print(f"File '{input_file_path}' not found.")
except Exception as e:
    print(f"An error occurred: {e}")
    
"""