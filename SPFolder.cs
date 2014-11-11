/*  Copyright (C) 2014 NAVERTICA a.s. http://www.navertica.com 

    This program is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License along
    with this program; if not, write to the Free Software Foundation, Inc.,
    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.  */
using System;
using System.Linq;
using System.IO;
using Microsoft.SharePoint;
using System.Collections.Generic;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;

namespace Navertica.SharePoint.Extensions
{
    public static class SPFolderExtensions
    {
        /// <summary>
        /// Checks if SPFolder with folderName exists in collection folders
        /// </summary>
        /// <param name="folders"></param>
        /// <param name="folderName"></param>
        /// <returns></returns>
        public static bool Contains(this SPFolderCollection folders, string folderName)
        {
            if (folders == null) throw new ArgumentNullException("folders");
            if (folderName == null) throw new ArgumentNullException("folderName");

            return folders.Cast<SPFolder>().Any(( f => f.Name == folderName ));
        }

        /// <summary>
        /// Checks if current folder Contains running workflow. It checks all items and folders to deep 
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        public static bool ContainsRunningWorkflows(this SPFolder folder)
        {
            if (folder == null) throw new ArgumentNullException("folder");

            //Check na samotnem folderu
            foreach (SPWorkflow wf in folder.Item.Workflows)
            {
                if (wf.InternalState != SPWorkflowState.Cancelled && !wf.IsCompleted)
                {
                    return true;
                }
            }

            bool result = false;

            folder.ProcessItems(delegate(SPListItem i)
            {
                if (i.IsFolder())
                {
                    result = ContainsRunningWorkflows(i.Folder);

                    if (result) throw new TerminateException();
                }
                else
                {
                    foreach (SPWorkflow wf in i.Workflows)
                    {
                        if (wf.InternalState != SPWorkflowState.Cancelled && !wf.IsCompleted)
                        {
                            result = true;
                            throw new TerminateException();
                        }
                    }
                }

                return null;
            });

            return result;
        }

        /// <summary>
        /// Copies entire folder including subfolders and items.
        /// </summary>
        /// <param name="folder">Folder to copy</param>
        /// <param name="toFolder">Target folder</param>
        /// <param name="deleteOriginal">True to delete original folder after successful copy</param>
        /// <param name="overwriteExisting">True to overwrite existing - if no queryStr is used, will look for the existing folder using filename in libraries and Title in lists</param>
        /// <param name="additional">Optional additional metadata fields to set in the copied folder - keys are field internal names</param>
        /// <param name="queryStr">Optional CAML query string to find existing folder to overwrite</param>
        /// <returns></returns>
        public static SPListItem CopyTo(this SPFolder folder, SPFolder toFolder, bool deleteOriginal = false, bool overwriteExisting = false, DictionaryNVR additional = null, string queryStr = "")
        {
            if (folder == null) throw new ArgumentNullException("folder");
            if (toFolder == null) throw new ArgumentNullException("toFolder");

            return folder.Item.CopyTo(toFolder, deleteOriginal, overwriteExisting, additional, queryStr);
        }

        /// <summary>
        /// Creates or Update a document from byte array in given folder
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="filename"></param>
        /// <param name="data"></param>
        /// <param name="overwrite"></param>
        /// <returns>newly created SPListItem</returns>
        public static SPListItem CreateOrUpdateDocument(this SPFolder folder, string filename, byte[] data, bool overwrite = false)
        {
            if (folder == null) throw new ArgumentNullException("folder");
            if (filename == null) throw new ArgumentNullException("filename");
            if (data == null) throw new ArgumentNullException("data");

            string newItemUrl = folder.ParentWeb.ServerRelativeUrl;

            if (newItemUrl.EndsWith("/"))
                newItemUrl += folder.Url + "/" + filename;
            else
                newItemUrl += "/" + folder.Url + "/" + filename;

            folder.Files.Add(newItemUrl, data, overwrite);
            SPListItem newItem = folder.ParentWeb.GetListItem(newItemUrl);

            return newItem;
        }

        /// <summary>
        /// For a path like "Documents/Folder1/Folder2", it either return the final folder, if it exists, or creates the folder 
        /// hierarchy and returns the final folder.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="path">illegal characters will be replaced with underscore</param>
        /// <returns></returns>
        public static SPFolder GetOrCreateFolder(this SPFolder folder, string path)
        {
            if (folder == null) throw new ArgumentNullException("folder");
            if (string.IsNullOrEmpty(path)) throw new ArgumentValidationException("path", "Cannot be null nor empty");

            // TODO path name normalization
            string[] folders = path.Split("/");
            SPFolder result = folder;
            bool first = true;

            folder.ParentWeb.RunWithAllowUnsafeUpdates(delegate
            {
                foreach (string originalFolderName in folders)
                {
                    if (string.IsNullOrEmpty(originalFolderName)) continue;
                    string folderName = originalFolderName.ReplaceInvalidFileNameChars('_');

                    SPFolder nextFolder = null;

                    if (result.Name == originalFolderName && first)
                    {
                        first = false;
                        continue;
                    }

                    foreach (SPFolder existing in result.SubFolders)
                    {
                        if (existing.Name == folderName)
                        {
                            nextFolder = existing;
                            break;
                        }
                    }

                    first = false;
                    if (nextFolder == null)
                    {
                        SPList list = folder.ParentWeb.OpenList(folder.ParentListId);
                        SPListItem newFolderItem = list.Items.Add(result.ServerRelativeUrl,
                            SPFileSystemObjectType.Folder, folderName);
                        newFolderItem["Title"] = originalFolderName;
                        if (list.ContainsFieldIntName("FileLeafRef")) newFolderItem["FileLeafRef"] = originalFolderName;
                        newFolderItem.Update();
                        nextFolder = newFolderItem.Folder;
                    }
                    result = nextFolder;
                }
            });

            return result;
        }

        /// <summary>
        /// Checks if the folder is the root folder
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        public static bool IsRootFolder(this SPFolder folder)
        {
            if (folder == null) throw new ArgumentNullException("folder");

            return folder.ParentFolder.ToString() == "";
        }

        /// <summary>
        /// Uploads file from a path into given SPFolder
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="directoryPath"></param>
        /// <param name="filename"></param>s
        /// <param name="overwrite">overwrite existing / add new version?</param>
        public static SPFile UploadFile(this SPFolder folder, string directoryPath, string filename, bool overwrite = false)
        {
            if (folder == null) throw new ArgumentNullException("folder");
            if (directoryPath == null) throw new ArgumentNullException("directoryPath");
            if (filename == null) throw new ArgumentNullException("directoryPath");

            using (FileStream fs = new FileStream(directoryPath + "\\" + filename, FileMode.Open, FileAccess.Read))
            {
                return folder.Files.Add(filename, fs.StreamToByteArray(), overwrite);
            }
        }

        /// <summary>
        /// Process folders, subfodlers and all items
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static ICollection<object> ProcessAllItems(this SPFolder folder, Func<SPListItem, object> func)
        {
            if (folder == null) throw new ArgumentNullException("folder");
            if (func == null) throw new ArgumentNullException("func");

            using (new SPMonitoredScope("ProcessAllItems - " + folder.Url))
            {
                List<object> result = new List<object>();

                folder.ProcessItems(delegate(SPListItem item)
                {
                    result.Add(item.ProcessItem(func));

                    if (item.IsFolder())
                    {
                        result.AddRange(ProcessAllItems(item.Folder, func));
                    }

                    return null;
                });

                return result;
            }
        }

        /// <summary>
        /// Process all items in folder using a delegate, can be also recursive
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="func"></param>
        /// <param name="includeSubFolderItems"></param>
        /// <returns></returns>
        public static ICollection<object> ProcessItems(this SPFolder folder, Func<SPListItem, object> func, bool includeSubFolderItems = false)
        {
            if (folder == null) throw new ArgumentNullException("folder");
            if (func == null) throw new ArgumentNullException("func");

            using (new SPMonitoredScope("ProcessItems - " + folder.Url))
            {
                SPQuery query = new SPQuery();
                query.Folder = folder;

                if (includeSubFolderItems)
                {
                    query.ViewAttributes = "Scope=\"Recursive\"";
                }

                return folder.ParentWeb.OpenList(folder.ParentListId).ProcessItems(func, query);
            }
        }

        /// <summary>
        /// Renames the folder - unsafe characters will be replaced with underscores
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="newName">forbidden characters will be replaced in folder name</param>
        public static void Rename(this SPFolder folder, string newName)
        {
            if (folder == null) throw new ArgumentNullException("folder");
            if (newName == null) throw new ArgumentNullException("newName");

            SPListItem item = folder.Item;
            item["Title"] = newName;
            item["BaseName"] = newName.ReplaceInvalidFileNameChars('_');
            item.SystemUpdate(false);
        }
    }
}