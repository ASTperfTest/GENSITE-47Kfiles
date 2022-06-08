using System.Data;
using Spring.Data;
using Gardening.Core.Domain;

namespace Gardening.Core.Persistence.ADO.RowMapper
{
    class TopicRowMapper : IRowMapper
    {
        public object MapRow(IDataReader dataReader, int rowNum)
        {
            Topic topic = new Topic();
            User owner = new User();
            owner.UserId = dataReader.GetString(dataReader.GetOrdinal("OWNER_ID"));
            owner.DisplayName = dataReader.GetString(dataReader.GetOrdinal("DISPLAY_NAME"));
            if (dataReader.IsDBNull(dataReader.GetOrdinal("NICKNAME")))
            {
                owner.Nickname = null;
            }
            else
            {
                owner.Nickname = dataReader.GetString(dataReader.GetOrdinal("NICKNAME"));
            }
            if (dataReader.IsDBNull(dataReader.GetOrdinal("EMAIL")))
            {
                owner.Email = null;
            }
            else
            {
                owner.Email = dataReader.GetString(dataReader.GetOrdinal("EMAIL"));
            }
            topic.Owner = owner;
            topic.Title = dataReader.GetString(dataReader.GetOrdinal("TITLE"));
            if (dataReader.IsDBNull(dataReader.GetOrdinal("DESCRIPTION")))
            {
                topic.Description = null;
            }
            else
            {
                topic.Description = dataReader.GetString(dataReader.GetOrdinal("DESCRIPTION"));
            }

            topic.TopicId = dataReader.GetString(dataReader.GetOrdinal("TOPIC_ID"));
            topic.IsApprove = dataReader.GetBoolean(dataReader.GetOrdinal("IS_APPROVE"));
            topic.CreateDateTime = dataReader.GetDateTime(dataReader.GetOrdinal("CREATE_DATETIME"));
            topic.ModifierId = dataReader.GetString(dataReader.GetOrdinal("MODIFIER_ID"));
            topic.ModifyDateTime = dataReader.GetDateTime(dataReader.GetOrdinal("MODIFY_DATETIME"));
            topic.RecommendedOrder = dataReader.GetInt32(dataReader.GetOrdinal("RECOMMENDED_ORDER"));
            return topic;
        }
    }
}
