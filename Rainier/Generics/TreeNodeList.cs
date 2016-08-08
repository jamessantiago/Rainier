using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rainier.Generics
{
	public class TreeNodeList<T> : List<TreeNode<T>>
	{
		public TreeNode<T> Parent;

		public TreeNodeList(TreeNode<T> Parent)
		{
			this.Parent = Parent;
		}

		public new TreeNode<T> Add(TreeNode<T> Node)
		{
			base.Add(Node);
			Node.Parent = Parent;
			return Node;
		}

		public TreeNode<T> Add(T Value)
		{
			TreeNode<T> Node = new TreeNode<T>(Parent);
			Node.Value = Value;
			return Node;
		}		
	}
}
