import os
import shutil

bulkFolderPath = "Path"
ArchiveFolderPath = "Path"

def find_pdf_with_prefix(folder_path, prefix):
    """
    Search for a PDF file in the specified folder with the given 8-character prefix.

    :param folder_path: Path of the folder to search in.
    :param prefix: The 8-character prefix to match.
    :return: The path of the first matching PDF file, or None if no match is found.
    """
    for file in os.listdir(folder_path):
        if file.startswith(prefix) and file.endswith('.pdf'):
            return os.path.join(folder_path, file)

    return None

for bulkItem in os.listdir(bulkFolderPath):
    if bulkItem.endswith(".pdf"):
        bulkItemName = bulkItem[:8]
        for archiveItem in os.listdir(ArchiveFolderPath):
            if archiveItem.startswith(bulkItemName):
                print(f"Found folder match for {bulkItemName}")
                archiveItemPath = os.path.join(ArchiveFolderPath, archiveItem)
                if find_pdf_with_prefix(archiveItemPath, bulkItemName) is None:
                    bulkItemPath = os.path.join(bulkFolderPath, bulkItem)
                    archiveItemPath = os.path.join(ArchiveFolderPath, archiveItem)
                    print(f"Copying {bulkItemPath} to {archiveItemPath}")
                    shutil.copy(bulkItemPath, archiveItemPath)
                    print(f"Copy complete")
                    break
                else:
                    print(f"PDF already exists for {bulkItemName}")
                    break
            else:
                print(f"No match found for {bulkItemName}")
    else:
        print(f"Skipping {bulkItem} because it is not a pdf file")
