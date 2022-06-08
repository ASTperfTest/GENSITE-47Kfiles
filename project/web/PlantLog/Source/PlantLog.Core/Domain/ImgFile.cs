using System;

namespace PlantLog.Core.Domain
{
    [Serializable]
    public class ImgFile
    {
        private string entryId;
        private string fileId;
        private string name;
        private string uri;

        public string EntryId
        {
            get
            {
                return entryId;
            }
            set
            {
                entryId = value;
            }
        }

        public string FileId
        {
            get
            {
                return fileId;
            }
            set
            {
                fileId = value;
            }
        }

        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        public string Uri
        {
            get
            {
                return uri;
            }
            set
            {
                uri = value;
            }
        }
    }
}
