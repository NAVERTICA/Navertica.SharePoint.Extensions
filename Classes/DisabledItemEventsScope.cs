using System;
using Microsoft.SharePoint;

namespace Navertica.SharePoint
{
    /// <summary>
    /// Use this scope to prevent event firing when updating SPListItems
    /// using (var scope = DisabledItemEventsScope()) { item.Update(); }
    /// via https://adrianhenke.wordpress.com/2010/01/29/disable-item-events-firing-during-item-update/
    /// </summary>
    public sealed class DisabledItemEventsScope : SPItemEventReceiver, IDisposable
    {
        private readonly bool oldValue;

        public bool Enabled
        {
            get { return base.EventFiringEnabled; }
        }

        public DisabledItemEventsScope()
        {
            oldValue = base.EventFiringEnabled;
            base.EventFiringEnabled = false;
        }

        public void Dispose()
        {
            base.EventFiringEnabled = oldValue;
        }
    }
}
