using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace Gardening.Core.Domain
{
    [Serializable]
    public class User
    {
        private string userId;
        private string displayName;
        private string nickname;
        private string email;

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
    }
}
