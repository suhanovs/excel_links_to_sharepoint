"""
Dennis Suhanovs
July 31, 2023

This script updates Excel links from local/network to SharePoint Online. 

Terminology: Model file --links_in--> Data file

How it works: provide the root of the local subtree for Excel search, and the 
root of the SharePoint library. The script searches the local subtree for all 
matching Excel files, opens them up, evaluates links, replaces them if necessary,
and saves (or saves-as) the Model file.

The SharePoint library must already contain the same subtree containing the 
Data files, so that when Model files are updated, the numeric values from Data 
files will continue to appear.

The script does NOT move any files to SharePoint.

Configurable elements: 
    - localFolderTree: root of the local subtree to search for workbooks with links
    - sharepointLibrary: root of the SharePoint library, to which you moved the
    "linked-to" workbooks. Example: https://<yourtenant>.sharepoint.com/sites/<yoursite>/<somefolder>
    - fileExtensions: a list of strings containing Excel file extensions to search for
    - linkFilter: a list of strings containing paths that will indicate which links
    need to be replaced (some models contain dead links that we don't need to touch)
"""
import xlwings as xw
from typing import List, Tuple
import os 
from tqdm import tqdm

""" constants """
xlLinkTypeExcelLinks = 1

""" user-configurable parameters """
sharepointLibrary = "https://<yourtenant>.sharepoint.com/sites/<yoursite>/<somefolder>"
localFolderTree = "c:\\users\\<username>\\documents"
fileExtensions = ['xlsx', 'xlsb']
linkFilter = ['N:\\', '\\\\<your_local_dfs_namespace\\', 'C:\\']



def breakup_path(path: str) -> Tuple[str,str]:
    filename = os.path.basename(path)
    pathname = os.path.dirname(path)
    return pathname, filename

def breakup_filename(filename: str) -> Tuple[str, str]:
    return os.path.splitext(filename)

def find_files(subtree: str, ext: List[str] = ['xlsx']) -> List:
    matches = []
    ext = [f".{x}" if not x.startswith('.') else x for x in ext]
    for dirpath, dirnames, filenames in os.walk(subtree):
        for filename in filenames:
            file, file_ext = breakup_filename(filename)
            if file_ext in ext:
                match = os.path.join(dirpath, filename)        
                matches.append(match)
    return matches

def update_links(path: str, linkFilter: List[str], saveas: bool = True):
    app = xw.App(visible=False)
    wb = app.books.open(path, update_links=None, read_only=False, ignore_read_only_recommended=True)
    links = wb.api.LinkSources()
    dirty = False
    try:
        if links:
            for link in links:
                if link.startswith(sharepointLibrary):
                    print(f"--- link {link} is already updated to point to SharePoint library")
                elif any(s in link for s in linkFilter):
                    newlink = f"{sharepointLibrary}/{breakup_path(link)[1]}"
                    print(f"--- found link {link}, updating to {newlink}")
                    wb.api.ChangeLink(
                        Name = link,
                        NewName = newlink,
                        Type = xlLinkTypeExcelLinks
                    )
                    dirty = True
                else:
                    print(f"--- found link {link} which does not match the link filter")

            if dirty:
                if saveas:
                    dir, filename = breakup_path(path)
                    name, ext = breakup_filename(filename)
                    newpath = f"{dir}/{name}_newlinks{ext}"
                    print(f"--- saving {path} as {newpath}")
                    wb.save(path=newpath)
                else:
                    print(f"--- saving in place: {path}")
                    wb.save()
            else:
                print(f"--- file {path} contained no links that required an update and was skipped")
        else:
            print(f"--- no links found in {path}")
    except Exception as ex:
        print(ex)
    finally:
        wb.close()
        app.quit()

def discover_links(path: str, linkFilter: List[str]):
    app = xw.App(visible=False)
    wb = app.books.open(path, update_links=None, read_only=False, ignore_read_only_recommended=True)
    links = wb.api.LinkSources()
    try:
        if links:
            for link in links:
                dir, file = breakup_path(link)
                newlink = f"{sharepointLibrary}/{file}"
                if link.startswith(sharepointLibrary):
                    print(f"--- link {link} is already updated to point to SharePoint library")
                elif any(s in link for s in linkFilter):
                    print(f"--- link {link} would be updated to {newlink}")
                else:
                    print(f"--- link {link} does not match the link filter and would be skipped")
        else:
            print("--- this workbook contains no links")
    except Exception as ex:
        print(ex)
    finally:
        wb.close()
        app.quit()



def main():
    excel_files = find_files(localFolderTree, ext=fileExtensions)

    """ investigate files only - no changes """
    for file in tqdm(excel_files): 
        print(f"\nInvestigating Excel file at path {file}")
        discover_links(file, linkFilter=linkFilter)

    """ update links but save as, do not change the models """
    #for file in tqdm(excel_files): 
    #    print(f"\nInvestigating Excel file at path {file}")
    #    update_links(file, linkFilter=linkFilter, saveas=True)

    """ update links and save in place - THINK BEFORE YOU DO """
    #for file in tqdm(excel_files): 
    #    print(f"\nInvestigating Excel file at path {file}")
    #    update_links(file, linkFilter=linkFilter, saveas=False)

if __name__ == '__main__':
    main()
