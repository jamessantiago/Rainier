using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rainier.Generics
{
	public class TreeNode<T> : IDisposable
	{
		#region Properties

		private TreeNode<T> _Parent;

		public TreeNode<T> Parent
		{
			get { return _Parent; }
			set
			{
				if (value == _Parent)
					return;

				if (_Parent != null)
					_Parent.Children.Remove(this);

				if (value != null && !value.Children.Contains(this))
					value.Children.Add(this);

				_Parent = value;
			}
		}

		public TreeNode<T> Root
		{
			get
			{
				TreeNode<T> node = this;
				while (node.Parent != null)
				{
					node = node.Parent;
				}
				return node;
			}
		}

		private TreeNodeList<T> _Children;

		public TreeNodeList<T> Children
		{
			get { return _Children; }
			private set { _Children = value; }
		}

		private T _Value;
		public T Value
		{
			get { return _Value; }
			set { _Value = value; }
		}

		private TreeTraversalDirection _DisposeTranversalDirection = TreeTraversalDirection.BottomUp;

		public TreeTraversalDirection DisposeTraversalDirection
		{
			get { return _DisposeTranversalDirection; }
			set { _DisposeTranversalDirection = value; }
		}

		private TreeTraversalType _DisposeTranversalType = TreeTraversalType.DepthFirst;

		public TreeTraversalType DisposeTraversalType
		{
			get { return _DisposeTranversalType; }
			set { _DisposeTranversalType = value; }
		}

		public int Depth
		{
			get
			{
				int depth = 0;
				TreeNode<T> node = this;
				while (node.Parent != null)
				{
					node = node.Parent;
					depth++;
				}
				return depth;
			}
		}

		#endregion Properties

		#region Constructor

		public TreeNode()
		{
			Parent = null;
			Children = new TreeNodeList<T>(this);
		}

		public TreeNode(T Value)
		{
			this.Value = Value;
			Children = new TreeNodeList<T>(this);
		}

		public TreeNode(TreeNode<T> Parent)
		{
			this.Parent = Parent;
			Children = new TreeNodeList<T>(this);
		}

		public TreeNode(TreeNodeList<T> Children)
		{
			Parent = null;
			this.Children = Children;
			Children.Parent = this;
		}

		public TreeNode(TreeNode<T> Parent, TreeNodeList<T> Children)
		{
			this.Parent = Parent;
			this.Children = Children;
			Children.Parent = this;
		}

		#endregion Constructor


		#region Functions

		public IEnumerable<TreeNode<T>> GetEnumerable(TreeTraversalType TraversalType, TreeTraversalDirection TraversalDirection)
		{
			switch (TraversalType)
			{
				case TreeTraversalType.DepthFirst: return GetDepthFirstEnumerable(TraversalDirection);
				case TreeTraversalType.BreadthFirst: return GetBreadthFirstEnumerable(TraversalDirection);
				default: return null;
			}
		}

		private IEnumerable<TreeNode<T>> GetDepthFirstEnumerable(TreeTraversalDirection TraversalDirection)
		{
			if (TraversalDirection == TreeTraversalDirection.TopDown)
				yield return this;

			foreach (TreeNode<T> child in Children)
			{
				var e = child.GetDepthFirstEnumerable(TraversalDirection).GetEnumerator();
				while (e.MoveNext())
				{
					yield return e.Current;
				}
			}

			if (TraversalDirection == TreeTraversalDirection.BottomUp)
				yield return this;
		}

		private IEnumerable<TreeNode<T>> GetBreadthFirstEnumerable(TreeTraversalDirection TraversalDirection)
		{
			if (TraversalDirection == TreeTraversalDirection.BottomUp)
			{
				var stack = new Stack<TreeNode<T>>();
				foreach(var item in GetBreadthFirstEnumerable(TreeTraversalDirection.TopDown))
				{
					stack.Push(item);
				}
				while (stack.Count > 0)
				{
					yield return stack.Pop();
				}
				yield break;
			}

			var queue = new Queue<TreeNode<T>>();
			queue.Enqueue(this);

			while(0 < queue.Count)
			{
				TreeNode<T> node = queue.Dequeue();

				foreach (TreeNode<T> child in node.Children)
				{
					queue.Enqueue(child);
				}

				yield return node;
			}
				
		}
		
		#endregion Functions

		#region Dispose

		public virtual void Dispose()
		{
			if (DisposeTraversalDirection == TreeTraversalDirection.BottomUp)
			{
				foreach (TreeNode<T> node in Children)
				{
					node.Dispose();
				}
			}

			if (DisposeTraversalDirection == TreeTraversalDirection.TopDown)
			{
				foreach (TreeNode<T> node in Children)
				{
					node.Dispose();
				}
			}
		}

		#endregion Dispose
	}
}
