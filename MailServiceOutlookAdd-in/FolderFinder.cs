using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailServiceOutlookAdd_in
{
    class FolderFinder
    {
        private readonly Outlook.Application _Application;
        public FolderFinder(Outlook.Application application)
        {
            _Application = application;
        }
        public Outlook.Folder FindFolderByPath(string selectedFolder)
        {
            string[] pathArr = PathToArray(selectedFolder);
            string id = GetFolderID(pathArr);

            Outlook.Folder folder = (Outlook.Folder)_Application.ActiveExplorer().Session.GetFolderFromID(id);
            return folder;
        }
        private string[] PathToArray(string path)
        {
            string fullPath = MailServiceSettings.RootProjectFolderName + "\\" + path;
            string[] strArrayOne = fullPath.Split('\\');
            return strArrayOne;
        }

        private string GetFolderID(string[] pathArr)
        {
            Outlook.Folder folder = (Outlook.Folder)_Application.ActiveExplorer().Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            int iterrator = 0;
            string id = GetFolder(folder, pathArr, iterrator);

            return id;
        }
        private string GetFolder(Outlook.Folder childFolder, string[] pathArr, int iterrator)
        {
            string cuurenFolderName = pathArr[iterrator];
            Outlook.Folder folder = (Outlook.Folder)childFolder.Folders[cuurenFolderName];

            if (pathArr.Length == ++iterrator)
            {
                return folder.EntryID;
            }
            return GetFolder(folder, pathArr, iterrator);
        }
    }
}
