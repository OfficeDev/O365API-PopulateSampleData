﻿//----------------------------------------------------------------------------------------------
//    Copyright 2015 Microsoft Corporation
//
//    Licensed under the MIT License (MIT);
//    you may not use this file except in compliance with the License.
//    You may obtain a copy of the License at
//
//      http://mit-license.org/
//
//    Unless required by applicable law or agreed to in writing, software
//    distributed under the License is distributed on an "AS IS" BASIS,
//    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//    See the License for the specific language governing permissions and
//    limitations under the License.
//----------------------------------------------------------------------------------------------

using Microsoft.Office365.SharePoint.CoreServices;
using Microsoft.Office365.SharePoint.FileServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.Pickers;

namespace O365DataApp.Helpers
{
    static class FileOperations
    {
        public static async Task addFiles(SharePointClient myFilesClient)
        {
            FileOpenPicker picker = new FileOpenPicker();
            picker.FileTypeFilter.Add("*");
            picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            IReadOnlyList<StorageFile> sFiles = await picker.PickMultipleFilesAsync();
            foreach (var sFile in sFiles)
            {
                if (sFile != null)
                {
                    using (var stream = await sFile.OpenStreamForReadAsync())
                    {
                        File newFile = new File
                        {
                            Name = sFile.Name
                        };
                        await myFilesClient.Files.AddItemAsync(newFile);
                        await myFilesClient.Files.GetById(newFile.Id).ToFile().UploadAsync(stream);
                    }
                }
            }
        }

        public static async Task<List<IItem>> getMyFiles(SharePointClient myFilesClient)
        {
            List<IItem> myFiles = new List<IItem>();
            var myFilesResult = await myFilesClient.Files.ExecuteAsync();
            do
            {
                var files = myFilesResult.CurrentPage;
                foreach (var myFile in files)
                {
                    myFiles.Add(myFile);
                }
                myFilesResult = await myFilesResult.GetNextPageAsync();
            } while (myFilesResult != null);
            return myFiles;
        }
    }
}
