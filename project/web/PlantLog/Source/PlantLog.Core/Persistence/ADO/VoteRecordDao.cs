using System;
using System.Data;
using PlantLog.Core.Domain;
using PlantLog.Core.Persistence.ADO.RowMapper;
using Spring.Data.Common;
using Spring.Data.Core;
using System.Collections;

namespace PlantLog.Core.Persistence.ADO
{
    public class VoteRecordDao : AdoDaoSupport, IVoteRecordDao
    {
        #region IVoteRecordDao Members

        public IList GetByUser(string userId)
        {
            if (userId.Trim() == "")
            {
                return null;
            }

            string cmd = "SELECT * FROM VOTE_RECORD WHERE USER_ID = @userId ORDER BY VOTE_DATE DESC";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("userId", DbType.String).Value = userId;

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd, new VoteRecordRowMapper(), dbParameters);
        }

        public IList GetVoteRecordByVoter(string voterId)
        {
            if (voterId.Trim() == "")
            {
                return null;
            }

            string cmd = "SELECT * FROM VOTE_RECORD WHERE VOTER_ID = @voterId ORDER BY VOTE_DATE DESC";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("voterId", DbType.String).Value = voterId;

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd, new VoteRecordRowMapper(), dbParameters);
        }

        public IList GetAll()
        {
            string cmd = "SELECT * FROM VOTE_RECORD ORDER BY USER_ID, VOTE_DATE DESC";

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd, new VoteRecordRowMapper());
        }

        public void Create(VoteRecord voteRecord)
        {
            if (voteRecord == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = @"INSERT INTO VOTE_RECORD (IP, VOTE_DATE, VOTER_ID, USER_ID) 
                        VALUES (@ip, @voteDate, @VoterId, @UserId)";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("ip", DbType.String).Value = voteRecord.IP;
            dbParameters.Add("voteDate", DbType.DateTime).Value = voteRecord.VoteDate;
            dbParameters.Add("VoterId", DbType.String).Value = voteRecord.VoterId;
            dbParameters.Add("UserId", DbType.String).Value = voteRecord.UserId;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        public void Update(VoteRecord voteRecord)
        {
            if (voteRecord == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = @"UPDATE VOTE_RECORD SET VOTE_DATE = @voteDate, USER_ID = @UserId, IP = @ip, 
                        VOTER_ID = @VoterId WHERE RECORD_ID = @id";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("id", DbType.Int32).Value = voteRecord.ID;
            dbParameters.Add("ip", DbType.String).Value = voteRecord.IP;
            dbParameters.Add("voteDate", DbType.DateTime).Value = voteRecord.VoteDate;
            dbParameters.Add("VoterId", DbType.String).Value = voteRecord.VoterId;
            dbParameters.Add("UserId", DbType.String).Value = voteRecord.UserId;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        #endregion
    }
}
