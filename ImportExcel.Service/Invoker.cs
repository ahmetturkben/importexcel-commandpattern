using ImportExcel.Service.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace ImportExcel.Service
{
    public class Invoker
    {
        #region Fields
        private ICommand _onStart { get; set; }

        private ICommand _onFinish { get; set; }
        #endregion

        #region Methods
        public void SetOnStart(ICommand command)
        {
            this._onStart = command;
        }

        public void SetOnFinish(ICommand command)
        {
            this._onFinish = command;
        }
        #endregion

        public void Running()
        {
            Console.WriteLine("Commmands running");

            if (this._onStart is ICommand)
                _onStart.Execute();

            if (this._onFinish is ICommand)
                _onFinish.Execute();

            Console.WriteLine("Commmands complated");
        }
    }
}
