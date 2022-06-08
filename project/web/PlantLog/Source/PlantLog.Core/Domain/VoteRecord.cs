using System;

namespace PlantLog.Core.Domain
{
    [Serializable]
    public class VoteRecord
    {
        private int id;
        private string ip;
        private DateTime voteDate;
        private string voterId;
        private string userId;

        public int ID
        {
            get
            {
                return id;
            }
            set
            {
                id = value;
            }
        }

        public string IP
        {
            get
            {
                return ip;
            }
            set
            {
                ip = value;
            }
        }

        public DateTime VoteDate
        {

            get
            {
                return voteDate;
            }
            set
            {
                voteDate = value;
            }
        }

        public string VoterId
        {
            get
            {
                return voterId;
            }
            set
            {
                voterId = value;
            }
        }

        public string UserId
        {
            get
            {
                return userId;
            }
            set
            {
                userId = value;
            }
        }
    }
}
