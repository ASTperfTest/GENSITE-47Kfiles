using System;
using System.Collections;
using System.Data;
using PlantLog.Core.Domain;
using Spring.Data.Common;
using Spring.Data.Core;
using PlantLog.Core.Persistence.ADO.RowMapper;

namespace PlantLog.Core.Persistence.ADO
{
    public class OwnerDao : AdoDaoSupport, IOwnerDao
    {
        #region IOwnerDao Members

        public Owner Get(string ownerId)
        {
            if (ownerId.Trim() == "")
            {
                return null;
            }

            if (!Exist(ownerId))
            {
                return null;
            }

            string cmd = "SELECT * FROM OWNER WHERE OWNER_ID = @ownerId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("ownerId", DbType.String).Value = ownerId;

            return (Owner)AdoTemplate.QueryForObject(CommandType.Text, cmd, new OwnerRowMapper(), dbParameters);
        }

        public IList GetAll()
        {
            string cmd = @"SELECT * FROM OWNER ORDER BY OWNER_ID";

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd, new OwnerRowMapper());
        }

        public Owner Create(Owner owner)
        {
            if (owner == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = @"INSERT INTO OWNER (OWNER_ID, DISPLAY_NAME, NICKNAME, TOPIC, EMAIL, [DESCRIPTION], CREATOR_ID, 
                        CREATE_DATETIME, MODIFIER_ID, MODIFY_DATETIME) VALUES (@OwnerId, @DisplayName, @Nickname, 
                        @Topic, @Email, @Description, @CreatorId, @CreateDateTime, @ModifierId, @ModifyDateTime)";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("OwnerId", DbType.String).Value = owner.OwnerId;
            dbParameters.Add("DisplayName", DbType.String).Value = owner.DisplayName;
            if (owner.Nickname == null)
            {
                dbParameters.Add("Nickname", DbType.String).Value = DBNull.Value;
            }
            else
            {
                dbParameters.Add("Nickname", DbType.String).Value = owner.Nickname;
            }
            dbParameters.Add("Topic", DbType.String).Value = owner.Topic;
            if (owner.Email == null)
            {
                dbParameters.Add("Email", DbType.String).Value = DBNull.Value;
            }
            else
            {
                dbParameters.Add("Email", DbType.String).Value = owner.Email;
            }
            if (owner.Description == null)
            {
                dbParameters.Add("Description", DbType.String).Value = DBNull.Value;
            }
            else
            {
                dbParameters.Add("Description", DbType.String).Value = owner.Description;
            }
            dbParameters.Add("CreatorId", DbType.String).Value = owner.CreatorId;
            dbParameters.Add("CreateDateTime", DbType.DateTime).Value = owner.CreateDateTime;
            dbParameters.Add("ModifierId", DbType.String).Value = owner.ModifierId;
            dbParameters.Add("ModifyDateTime", DbType.DateTime).Value = owner.ModifyDateTime;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);

            return Get(owner.OwnerId);
        }

        public Owner Update(Owner owner)
        {
            if (owner == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = @"UPDATE OWNER SET TOPIC = @Topic, [DESCRIPTION] = @Description, 
                        IS_APPROVE = @IsApprove,EMAIL = @email, MODIFIER_ID = @ModifierId, 
                        MODIFY_DATETIME = @ModifyDateTime WHERE OWNER_ID = @OwnerId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("OwnerId", DbType.String).Value = owner.OwnerId;
            dbParameters.Add("Topic", DbType.String).Value = owner.Topic;
            if (owner.Email == null)
            {
                dbParameters.Add("email", DbType.String).Value = DBNull.Value;
            }
            else
            {
                dbParameters.Add("email", DbType.String).Value = owner.Email;
            }
            if (owner.Description == null)
            {
                dbParameters.Add("Description", DbType.String).Value = DBNull.Value;
            }
            else
            {
                dbParameters.Add("Description", DbType.String).Value = owner.Description;
            }
            dbParameters.Add("IsApprove", DbType.Boolean).Value = owner.IsApprove;
            dbParameters.Add("ModifierId", DbType.String).Value = owner.ModifierId;
            dbParameters.Add("ModifyDateTime", DbType.DateTime).Value = owner.ModifyDateTime;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);

            return Get(owner.OwnerId);
        }

        public void Delete(string ownerId)
        {
            string cmd = "DELETE FROM OWNER WHERE OWNER_ID = @OwnerId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("OwnerId", DbType.String).Value = ownerId;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        #endregion

        private bool Exist(string ownerId)
        {
            string cmd = "SELECT * FROM OWNER WHERE OWNER_ID = @ownerId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("ownerId", DbType.String).Value = ownerId;

            DataTable result = AdoTemplate.DataTableCreateWithParams(CommandType.Text, cmd, dbParameters);

            if (result.Rows.Count == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
