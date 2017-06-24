using System.Linq;

namespace VBASync.Model.FrxObjects
{
    internal class RawControl
    {
        private readonly byte[] _content;

        internal RawControl(byte[] content)
        {
            _content = content;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
            {
                return false;
            }
            if (ReferenceEquals(this, obj))
            {
                return true;
            }
            if (obj.GetType() != GetType())
            {
                return false;
            }
            return Equals((RawControl)obj);
        }

        public override int GetHashCode()
        {
            return _content?.Length.GetHashCode() ?? 0;
        }

        protected bool Equals(RawControl other)
        {
            return _content != null && other._content != null && _content.SequenceEqual(other._content);
        }
    }
}
