﻿using System;
using System.Collections.Generic;

namespace Subtitution
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Manager oManager = new Manager();
            System.Windows.Forms.Application.Run();
        }
    }
}
