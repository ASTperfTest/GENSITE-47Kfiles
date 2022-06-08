using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Spring.Data;
using System.Data;
using PlantLog.Core.Domain;

namespace PlantLog.Core.Persistence.ADO.RowMapper
{
    class VoteRecordRowMapper : IRowMapper
    {
        public object MapRow(IDataReader dataReader, int rowNum)
        {
            VoteRecord voteRecord = new VoteRecord();
            voteRecord.ID = dataReader.GetInt32(0);
            voteRecord.IP = dataReader.GetString(1);
            voteRecord.VoteDate = dataReader.GetDateTime(2);
            voteRecord.VoterId = dataReader.GetString(3);
            voteRecord.UserId = dataReader.GetString(4);

            return voteRecord;
        }
    }
}
