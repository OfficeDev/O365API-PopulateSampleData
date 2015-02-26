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
