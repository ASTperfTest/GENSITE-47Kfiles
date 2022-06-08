using System;
using System.Collections;
using System.Xml.Serialization;

namespace PlantLog.Core.Domain
{
    [XmlInclude(typeof(ImgFile))]
    [Serializable]
    public class Entry
    {
        private string entryId;
        private string ownerId;
        private DateTime date;
        private string title;
        private string description;
        private IList files=new ArrayList();
        private bool isPublic;
        private bool isApprove;
        private string creatorId;
        private DateTime createDateTime;
        private string modifierId;
        private DateTime modifyDateTime;

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

        public string OwnerId
        {
            get
            {
                return ownerId;
            }
            set
            {
                ownerId = value;
            }
        }

        public DateTime Date
        {
            get
            {
                return date;
            }
            set
            {
                date = value;
            }
        }

        public string Title
        {
            get
            {
                return title;
            }
            set
            {
                title = value;
            }
        }

        public string Description
        {
            get
            {
                return description;
            }
            set
            {
                description = value;
            }
        }

        public IList Files
        {
            get
            {
                return files;
            }
            set
            {
                files = value;
            }
        }

        public bool IsPublic
        {
            get
            {
                return isPublic;
            }
            set
            {
                isPublic = value;
            }
        }

        public bool IsApprove
        {
            get
            {
                return isApprove;
            }
            set
            {
                isApprove = value;
            }
        }

        public string CreatorId
        {
            get
            {
                return creatorId;
            }
            set
            {
                creatorId = value;
            }
        }

        public DateTime CreateDateTime
        {
            get
            {
                return createDateTime;
            }
            set
            {
                createDateTime = value;
            }
        }

        public string ModifierId
        {
            get
            {
                return modifierId;
            }
            set
            {
                modifierId = value;
            }
        }

        public DateTime ModifyDateTime
        {
            get
            {
                return modifyDateTime;
            }
            set
            {
                modifyDateTime = value;
            }
        }
    }
}
