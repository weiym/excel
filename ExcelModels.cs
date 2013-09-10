using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Excel
{
    class ExcelModels
    {
        #region 属性
        private String id;//序号
        private String workflowId;//工作流ID
        private String workflowName;//工作流名称
        private String workflowStepId;//环节ID
        private String workflowStepName;//环节名称
        private String workflowStepPath;//可提交路径
        private String workflowOPinion;//审批意见是否必填
        private String workflowSwitches;//开关编码
        private String workflowNote;//备注

        

        #endregion


        #region 属性

        /// <summary>
        /// 序号
        /// </summary>
        public String Id
        {
            get { return id; }
            set { id = value; }
        }
        
        /// <summary>
        /// 工作流ID
        /// </summary>
        public String WorkflowId
        {
            get { return workflowId; }
            set { workflowId = value; }
        }
        
        /// <summary>
        /// 工作流名称
        /// </summary>
        public String WorkflowName
        {
            get { return workflowName; }
            set { workflowName = value; }
        }
        
        /// <summary>
        /// 环节ID
        /// </summary>
        public String WorkflowStepId
        {
            get { return workflowStepId; }
            set { workflowStepId = value; }
        }
        
        /// <summary>
        /// 环节名称
        /// </summary>
        public String WorkflowStepName
        {
            get { return workflowStepName; }
            set { workflowStepName = value; }
        }
        
        /// <summary>
        /// 可提交路径
        /// </summary>
        public String WorkflowStepPath
        {
            get { return workflowStepPath; }
            set { workflowStepPath = value; }
        }
        
        /// <summary>
        /// 审批意见是否必填
        /// </summary>
        public String WorkflowOPinion
        {
            get { return workflowOPinion; }
            set { workflowOPinion = value; }
        }
        
        /// <summary>
        /// 开关编码
        /// </summary>
        public String WorkflowSwitches
        {
            get { return workflowSwitches; }
            set { workflowSwitches = value; }
        }
        /// <summary>
        /// 备注
        /// </summary>
        public String WorkflowNote
        {
            get { return workflowNote; }
            set { workflowNote = value; }
        }

        #endregion

    }
}
