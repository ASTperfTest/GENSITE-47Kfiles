using System;
using System.Collections;
using System.Xml.Serialization;

namespace Gardening.Core.Domain
{
    [Serializable]
    public class Topic
    {
        private string topicId;
        private ImgFile avatar;
        private string title;
        private string description;
        private bool isApprove;
        private User owner; //owner為creator
        private DateTime createDateTime;
        private string modifierId;
        private DateTime modifyDateTime;
        private int recommendedOrder; //首頁推薦Topic順序 
        
        public string TopicId
        {
            get
            {
                return topicId;
            }
            set
            {
                topicId = value;
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

        public User Owner
        {
            get
            {
                return owner;
            }
            set
            {
                owner = value;
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

        public int RecommendedOrder
        {
            get
            {
                return recommendedOrder;
            }
            set
            {
                recommendedOrder = value;
            }
        }
    }
}
