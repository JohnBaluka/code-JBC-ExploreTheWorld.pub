using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.CL
{
    public class TreeNode
    {
        public TreeNode()
        {
        }
        
        public string ParentID { get; set; } = string.Empty;
        public string ID { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
        public bool Checked { get; set; }
        public bool Expanded { get; set; }

        public List<TreeNode> Children { get; set; } = new List<TreeNode>();

    }
}
