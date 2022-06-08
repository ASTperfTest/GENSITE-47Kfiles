using System;
using System.Collections;
using System.Data;
using Gardening.Core.Domain;
using Spring.Data.Common;
using Spring.Data.Core;
using Gardening.Core.Persistence.ADO.RowMapper;
using System.Text;

namespace Gardening.Core.Persistence.ADO
{
    public class TopicDao : AdoDaoSupport, ITopicDao
    {
        #region IOwnerDao Members

        public Topic Get(string topicId)
        {
            if (topicId.Trim() == "")
            {
                return null;
            }

            if (!topicExist(topicId))
            {
                return null;
            }

            StringBuilder cmd = new StringBuilder();
            cmd.Append(" SELECT T.*,U.DISPLAY_NAME,U.NICKNAME,U.EMAIL " + Environment.NewLine);
            cmd.Append(" FROM TOPIC T " + Environment.NewLine);
            cmd.Append(" LEFT JOIN [USER] U " + Environment.NewLine);
            cmd.Append(" ON T.OWNER_ID=U.USER_ID " + Environment.NewLine);
            cmd.Append(" WHERE T.TOPIC_ID=@TopicId " + Environment.NewLine);
            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("TopicId", DbType.String).Value = topicId;

            return (Topic)AdoTemplate.QueryForObject(CommandType.Text, cmd.ToString(), new TopicRowMapper(), dbParameters);
        }

        public IList GetAll()
        {
            StringBuilder cmd = new StringBuilder();
            cmd.Append(" SELECT T.*,U.DISPLAY_NAME,U.NICKNAME,U.EMAIL " + Environment.NewLine);
            cmd.Append(" FROM TOPIC T " + Environment.NewLine);
            cmd.Append(" LEFT JOIN [USER] U " + Environment.NewLine);
            cmd.Append(" ON T.OWNER_ID=U.USER_ID " + Environment.NewLine);
            cmd.Append(" ORDER BY T.CREATE_DATETIME DESC " + Environment.NewLine);

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd.ToString(), new TopicRowMapper());
        }

        public IList GetByOwner(string ownerId)
        {
            StringBuilder cmd = new StringBuilder();
            cmd.Append(" SELECT T.*,U.DISPLAY_NAME,U.NICKNAME,U.EMAIL " + Environment.NewLine);
            cmd.Append(" FROM TOPIC T " + Environment.NewLine);
            cmd.Append(" LEFT JOIN [USER] U " + Environment.NewLine);
            cmd.Append(" ON T.OWNER_ID=U.USER_ID " + Environment.NewLine);
            cmd.Append(" WHERE OWNER_ID=@OwnerId ORDER BY T.CREATE_DATETIME DESC " + Environment.NewLine);

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("OwnerId", DbType.String).Value = ownerId;

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd.ToString(), new TopicRowMapper(), dbParameters);
        }

        public Topic Create(Topic topic)
        {
            if (topic == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = @"INSERT INTO TOPIC (TOPIC_ID, TITLE, [DESCRIPTION], OWNER_ID, 
                        CREATE_DATETIME, MODIFIER_ID, MODIFY_DATETIME, RECOMMENDED_ORDER) VALUES (@TopicId, 
                        @Title, @Description, @OwnerId, @CreateDateTime, @ModifierId, @ModifyDateTime, @RecommendedOrder)";

            IDbParameters dbParameters = CreateDbParameters();
            topic.TopicId = Utility.GetGuid();
            dbParameters.Add("TopicId", DbType.String).Value = topic.TopicId;
            dbParameters.Add("OwnerId", DbType.String).Value = topic.Owner.UserId;
            dbParameters.Add("Title", DbType.String).Value = topic.Title;
            dbParameters.Add("CreateDateTime", DbType.DateTime).Value = topic.CreateDateTime;
            dbParameters.Add("ModifierId", DbType.String).Value = topic.ModifierId;
            dbParameters.Add("ModifyDateTime", DbType.DateTime).Value = topic.ModifyDateTime;
            dbParameters.Add("RecommendedOrder", DbType.Int32).Value = topic.RecommendedOrder;

            if (topic.Description == null)
            {
                dbParameters.Add("Description", DbType.String).Value = DBNull.Value;
            }
            else
            {
                dbParameters.Add("Description", DbType.String).Value = topic.Description;
            }

            if (!UserExist(topic.Owner.UserId))
            {
                CreateUser(topic.Owner);
            }

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);

            return Get(topic.TopicId);
        }

        public Topic Update(Topic topic)
        {
            if (topic == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = @"UPDATE TOPIC SET TITLE = @Title, [DESCRIPTION] = @Description, 
                        IS_APPROVE = @IsApprove, MODIFIER_ID = @ModifierId, 
                        MODIFY_DATETIME = @ModifyDateTime, RECOMMENDED_ORDER= @RecommendedOrder WHERE TOPIC_ID = @TopicId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("TopicId", DbType.String).Value = topic.TopicId;
            dbParameters.Add("Title", DbType.String).Value = topic.Title;
            if (topic.Description == null)
            {
                dbParameters.Add("Description", DbType.String).Value = DBNull.Value;
            }
            else
            {
                dbParameters.Add("Description", DbType.String).Value = topic.Description;
            }
            dbParameters.Add("IsApprove", DbType.Boolean).Value = topic.IsApprove;
            dbParameters.Add("ModifierId", DbType.String).Value = topic.ModifierId;
            dbParameters.Add("ModifyDateTime", DbType.DateTime).Value = topic.ModifyDateTime;
            dbParameters.Add("RecommendedOrder", DbType.Int32).Value = topic.RecommendedOrder;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);

            return Get(topic.TopicId);
        }

        public void Delete(string topicId)
        {
            string cmd = "DELETE FROM TOPIC WHERE TOPIC_ID = @TopicId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("TopicId", DbType.String).Value = topicId;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        public IList GetByRecommendedOrder(int topN)
        {
            StringBuilder cmd = new StringBuilder();
            cmd.Append(" SELECT TOP " + topN.ToString() + " T.*,U.DISPLAY_NAME,U.NICKNAME,U.EMAIL " + Environment.NewLine);
            cmd.Append(" FROM TOPIC T " + Environment.NewLine);
            cmd.Append(" LEFT JOIN [USER] U " + Environment.NewLine);
            cmd.Append(" ON T.OWNER_ID=U.USER_ID " + Environment.NewLine);
            cmd.Append(" ORDER BY T.RECOMMENDED_ORDER " + Environment.NewLine);

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd.ToString(), new TopicRowMapper());
        }

        #endregion

        private bool UserExist(string userId)
        {
            string cmd = "SELECT * FROM [USER] WHERE USER_ID=@UserId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("UserId", DbType.String).Value = userId;

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

        private void CreateUser(User user)
        {
            if (user == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = @"INSERT INTO [USER] (USER_ID, DISPLAY_NAME, NICKNAME, EMAIL 
                        ) VALUES (@UserId, @DisplayName, @Nickname, @Email)";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("UserId", DbType.String).Value = user.UserId;
            dbParameters.Add("DisplayName", DbType.String).Value = user.DisplayName;

            if (user.Nickname == null)
            {
                dbParameters.Add("Nickname", DbType.String).Value = DBNull.Value;
            }
            else
            {
                dbParameters.Add("Nickname", DbType.String).Value = user.Nickname;
            }

            if (user.Email == null)
            {
                dbParameters.Add("Email", DbType.String).Value = DBNull.Value;
            }
            else
            {
                dbParameters.Add("Email", DbType.String).Value = user.Email;
            }

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        private bool topicExist(string topicId)
        {
            string cmd = "SELECT * FROM TOPIC WHERE TOPIC_ID = @TopicId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("TopicId", DbType.String).Value = topicId;

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

        public IList GetSearchResult(string keyword, int pageIndex, int pageSize)
        {
            StringBuilder cmd = new StringBuilder();
            cmd.Append(" select t.*,U.DISPLAY_NAME,U.NICKNAME,U.EMAIL " + Environment.NewLine);
            cmd.Append(" FROM ( " + Environment.NewLine);
            cmd.Append("       SELECT rank() OVER ( ORDER BY MODIFY_DATETIME DESC ) AS RankNumber,* FROM TOPIC" + Environment.NewLine);
            cmd.Append("       where (TITLE like @Keywoed OR DESCRIPTION like @Keywoed or TOPIC_ID IN " + Environment.NewLine);
            cmd.Append("       (SELECT ENTRY.TOPIC_ID From ENTRY  " + Environment.NewLine);
            cmd.Append("            where (ENTRY.TITLE like @Keywoed OR ENTRY.DESCRIPTION like @Keywoed) and ENTRY.IS_PUBLIC = '1' " + Environment.NewLine);
            cmd.Append("       Group by ENTRY.TOPIC_ID )" + Environment.NewLine);
            cmd.Append("       )" + Environment.NewLine);
            cmd.Append(")as t LEFT JOIN [USER] U ON T.OWNER_ID=U.USER_ID where RankNumber between " + (pageSize * (pageIndex)+1).ToString() + " and " + (pageSize * (pageIndex + 1)).ToString() + Environment.NewLine);
            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("Keywoed", DbType.String).Value = '%' + keyword + '%';

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd.ToString(), new TopicRowMapper(), dbParameters);
        }
    }
}
