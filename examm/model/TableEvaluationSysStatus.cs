namespace examm.model
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class TableEvaluationSysStatus
    {
        [Key]
        public int intEvaluationSysStatusId { get; set; }

        [StringLength(250)]
        public string strStatysName { get; set; }
    }
}
