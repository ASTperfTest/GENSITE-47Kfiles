using System;
using System.Collections;
using System.Data;
using Gardening.Core.Domain;
using Gardening.Core.Persistence.ADO.RowMapper;
using Spring.Data.Common;
using Spring.Data.Core;

namespace Gardening.Core.Persistence.ADO
{
    public class EntryDao : AdoDaoSupport, IEntryDao
    {
        #region IEntryDao Members

        public Entry Get(string entryId)
        {
            if (entryId.Trim() == "")
            {
                return null;
            }

            string cmd = "SELECT * FROM ENTRY WHERE ENTRY_ID = @entryId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("entryId", DbType.String).Value = entryId;

            return (Entry)AdoTemplate.QueryForObject(CommandType.Text, cmd, new EntryRowMapper(), dbParameters);
        }

        public IList GetByTopic(string topicId)
        {
            if (topicId.Trim() == "")
            {
                return null;
            }

            string cmd = "SELECT * FROM ENTRY WHERE TOPIC_ID = @TopicId ORDER BY DATE , CREATE_DATETIME ";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("TopicId", DbType.String).Value = topicId;

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd, new EntryRowMapper(), dbParameters);
        }

        public Entry Create(Entry entry)
        {
            if (entry == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = @"INSERT INTO ENTRY (ENTRY_ID, TOPIC_ID, [DATE], TITLE, [DESCRIPTION], IS_PUBLIC, CREATOR_ID, 
                        CREATE_DATETIME, MODIFIER_ID, MODIFY_DATETIME) VALUES (@EntryId, @TopicId, @Date, @Title, 
                        @Description, @IsPublic, @CreatorId, @CreateDateTime, @ModifierId, @ModifyDateTime)";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("EntryId", DbType.String).Value = entry.EntryId;
            dbParameters.Add("TopicId", DbType.String).Value = entry.TopicId;
            dbParameters.Add("Date", DbType.DateTime).Value = entry.Date;
            dbParameters.Add("Title", DbType.String).Value = entry.Title;
            dbParameters.Add("Description", DbType.String).Value = entry.Description;
            dbParameters.Add("IsPublic", DbType.Boolean).Value = entry.IsPublic;
            dbParameters.Add("CreatorId", DbType.String).Value = entry.CreatorId;
            dbParameters.Add("CreateDateTime", DbType.DateTime).Value = entry.CreateDateTime;
            dbParameters.Add("ModifierId", DbType.String).Value = entry.ModifierId;
            dbParameters.Add("ModifyDateTime", DbType.DateTime).Value = entry.ModifyDateTime;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);

            return Get(entry.EntryId);
        }

        public Entry Update(Entry entry)
        {
            if (entry == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = @"UPDATE ENTRY SET [DATE] = @Date, TITLE = @Title, [DESCRIPTION] = @Description, 
                        IS_PUBLIC = @IsPublic, IS_APPROVE = @IsApprove, MODIFIER_ID = @ModifierId, 
                        MODIFY_DATETIME = @ModifyDateTime  WHERE ENTRY_ID = @EntryId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("EntryId", DbType.String).Value = entry.EntryId;
            dbParameters.Add("Date", DbType.DateTime).Value = entry.Date;
            dbParameters.Add("Title", DbType.String).Value = entry.Title;
            dbParameters.Add("Description", DbType.String).Value = entry.Description;
            dbParameters.Add("IsPublic", DbType.Boolean).Value = entry.IsPublic;
            dbParameters.Add("IsApprove", DbType.Boolean).Value = entry.IsApprove;
            dbParameters.Add("ModifierId", DbType.String).Value = entry.ModifierId;
            dbParameters.Add("ModifyDateTime", DbType.DateTime).Value = entry.ModifyDateTime;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);

            return Get(entry.EntryId);
        }

        public void Delete(string entryId)
        {
            string cmd = "DELETE FROM ENTRY WHERE ENTRY_ID = @EntryId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("EntryId", DbType.String).Value = entryId;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        public void DeleteByTopic(string topicId)
        {
            string cmd = "DELETE FROM ENTRY WHERE TOPIC_ID = @TopicId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("TopicId", DbType.String).Value = topicId;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        public IList GetAll()
        {
            string cmd = "SELECT * FROM ENTRY";
            IDbParameters dbParameters = CreateDbParameters();
            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd, new EntryRowMapper(), dbParameters);
        }

        #endregion
    }
}
