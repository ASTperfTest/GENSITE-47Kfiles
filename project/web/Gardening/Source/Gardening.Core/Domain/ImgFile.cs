using System;

namespace Gardening.Core.Domain
{
    [Serializable]
    public class ImgFile
    {
        private string parentId;
        private string fileId;
        private string name;
        private string uri;

        public string ParentId
        {
            get
            {
                return parentId;
            }
            set
            {
                parentId = value;
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
