/*  Copyright (C) 2014 NAVERTICA a.s. http://www.navertica.com 

    This program is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License along
    with this program; if not, write to the Free Software Foundation, Inc.,
    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.  */

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
