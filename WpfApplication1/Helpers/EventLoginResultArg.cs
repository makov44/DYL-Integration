using System;

namespace DYL.EmailIntegration.Helpers
{
    public class EventLoginResultArg: EventArgs
    {
        public LoginResult LoginResult { set; get; }

        public EventLoginResultArg(LoginResult loginResul)
        {
            LoginResult = loginResul;
        }
    }
}