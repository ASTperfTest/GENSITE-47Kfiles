using System.Data;
using Gardening.Core.Domain;
using Spring.Data;

namespace Gardening.Core.Persistence.ADO.RowMapper
{
    class EntryRowMapper : IRowMapper
    {
        public object MapRow(IDataReader dataReader, int rowNum)
        {
            Entry entry = new Entry();
            entry.EntryId = dataReader.GetString(0);
            entry.TopicId = dataReader.GetString(1);
            entry.Date = dataReader.GetDateTime(2);
            entry.Title = dataReader.GetString(3);
            entry.Description = dataReader.GetString(4);
            entry.IsPublic = dataReader.GetBoolean(5);
            entry.IsApprove = dataReader.GetBoolean(6);
            entry.CreatorId = dataReader.GetString(7);
            entry.CreateDateTime = dataReader.GetDateTime(8);
            entry.ModifierId = dataReader.GetString(9);
            entry.ModifyDateTime = dataReader.GetDateTime(10);

            return entry;
        }
    }
}
