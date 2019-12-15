using System;
using System.Collections.Generic;

namespace app
{
    public class SahamyabMarketInfo
    {
        public SahamyabMarketInfo()
        {


        }
        public List<Result> result { get; set; }
        public bool success { get; set; }

    }

    public class Result
    {
        public float index_affect { get; set; }
        public float index_affect_rank { get; set; }
        public float PE { get; set; }
        public float sectorPE { get; set; }
        public float profit7Days { get; set; }
        public float profit30Days { get; set; }
        public float profit91Days { get; set; }
        public float profit182Days { get; set; }
        public float profit365Days { get; set; }
        public float profitAllDays { get; set; }
        public float monthProfitRank { get; set; }
        public float monthProfitRankGroup { get; set; }
        public float marketValueRank { get; set; }
        public float marketValueRankGroup { get; set; }
        public float tradeVolumeRank { get; set; }
        public float tradeVolumeRankGroup { get; set; }
        public float zaribNaghdShavandegi { get; set; }
        public float correlation_dollar { get; set; }
        public float correlation_ons_tala { get; set; }
        public float correlation_oil_opec { get; set; }
        public float correlation_main_index { get; set; }
        public float sahamayb_post_count { get; set; }
        public float sahamayb_post_count_rank { get; set; }
        public float sahamyab_follower_count_rank { get; set; }
        public float sahamyab_page_visit_rank { get; set; }
    }
}
