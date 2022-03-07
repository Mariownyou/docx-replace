import zipfile
import os
from os.path import basename

def zip_dir(directory, zipname):
    """
    Compress a directory (ZIP file).
    """
    if os.path.exists(directory):
        outZipFile = zipfile.ZipFile(zipname, 'w', zipfile.ZIP_DEFLATED)

        # The root directory within the ZIP file.
        rootdir = os.path.basename(directory)

        for dirpath, dirnames, filenames in os.walk(directory):
            for filename in filenames:

                # Write the file named filename to the archive,
                # giving it the archive name 'arcname'.
                filepath   = os.path.join(dirpath, filename)
                parentpath = os.path.relpath(filepath, directory)
                arcname    = os.path.join(rootdir, parentpath)
                print(filepath)
                print(arcname)
                print()

                outZipFile.write(filepath, arcname)

    outZipFile.close()

zip_dir('output/', 'doc.docx')