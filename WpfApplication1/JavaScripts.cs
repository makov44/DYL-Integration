using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApplication1
{
    static class JavaScripts
    {
        public static string AlertBlocker = @"window.alert = function () { return true;  }; window.confirm=function () { return true; }; window.close=function () { };";
        public static string OnbeforeunloadBlocker = @"(function () {
            var onbeforeunloadHandler = function (ev) {
                if (ev) {
                    if (ev.stopPropagation)
                        ev.stopPropagation();
                    if (ev.stopImmediatePropagation)
                        ev.stopImmediatePropagation();
                    ev.returnValue = undefined;
                }
                window.event.returnValue = undefined;
            }

            var handler = null;
            var intervalHandler = function () {
                if (handler)
                    window.detachEvent('onbeforeunload', handler);

            // window.attachEvent works best
            handler = window.attachEvent('onbeforeunload', onbeforeunloadHandler);
            // handler = window.addEventListener('beforeunload', onbeforeunloadHandler);
            // handler = window.onload = onbeforeunloadHandler;
        };

        window.setInterval(intervalHandler, 500);
        intervalHandler();
    })();";
    }
}
