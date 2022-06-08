using System;
using System.Collections;
using System.Xml.Serialization;

namespace PlantLog.Core.Domain
{
    [XmlInclude(typeof(VoteRecord))]
    [Serializable]
    public class Owner
    {
        private string ownerId;
        private string displayName;
        private string nickname;
        private ImgFile avatar;
        private string topic;
        private string description;
        private IList voteRecords;
        private string email;
        private bool isApprove;
        private string creatorId;
        private DateTime createDateTime;
        private string modifierId;
        private DateTime modifyDateTime;

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

        public string DisplayName
        {
            get
            {
                return displayName;
            }
            set
            {
                displayName = value;
            }
        }

        public string Nickname
        {
            get
            {
                return nickname;
            }
            set
            {
                nickname = value;
            }
        }

        public ImgFile Avatar
        {
            get
            {
                return avatar;
            }
            set
            {
                avatar = value;
            }
        }

        public string Topic
        {
            get
            {
                return topic;
            }
            set
            {
                topic = value;
            }
        }

        public string Email
        {
            get
            {
                return email;
            }
            set
            {
                email = value;
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

        public IList VoteRecords
        {
            get
            {
                return voteRecords;
            }
            set
            {
                voteRecords = value;
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
