using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace DYL.EmailIntegration.Helpers
{
    public class ObservableConcurrentBag<T> : ConcurrentQueue<T>, INotifyCollectionChanged
    {
        public event NotifyCollectionChangedEventHandler CollectionChanged;

        /// <summary>Adds an object to the end of the <see cref="T:System.Collections.Concurrent.ConcurrentQueue`1" />.</summary>
        /// <param name="item">The object to add to the end of the <see cref="T:System.Collections.Concurrent.ConcurrentQueue`1" />. The value can be a null reference (Nothing in Visual Basic) for reference types.</param>
        public new void Enqueue(T item)
        {
            OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add, item, 0));
            base.Enqueue(item);
        }

        /// <summary>Tries to remove and return the object at the beginning of the concurrent queue.</summary>
        /// <returns>true if an element was removed and returned from the beginning of the <see cref="T:System.Collections.Concurrent.ConcurrentQueue`1" /> successfully; otherwise, false.</returns>
        /// <param name="result">When this method returns, if the operation was successful, <paramref name="result" /> contains the object removed. If no object was available to be removed, the value is unspecified.</param>
        public new bool TryDequeue(out T result)
        {

            var isSuccess =  base.TryDequeue(out result);
            OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Remove, result, 0));
            return isSuccess;
        }
      
        private  void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
        {
            CollectionChanged?.Invoke(this, e);
        }
    }
}
