﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DYL.EmailIntegration.Domain.Data;
using log4net.Config;

namespace DYL.EmailIntegration.Helpers
{
    public class EventSessionArgs : EventArgs
    {
        public Credentials Credentials { get; set; }

        public EventSessionArgs(Credentials credentials)
        {
            Credentials = credentials;
        }
    }
}
