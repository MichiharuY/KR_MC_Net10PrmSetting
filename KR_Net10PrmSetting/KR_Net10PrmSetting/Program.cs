namespace KR_Net10PrmSetting
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            FormSplash fs = new FormSplash();
            fs.StartPosition = FormStartPosition.CenterScreen;
            fs.Show();
            fs.Refresh();
            Thread.Sleep(2000);//���Ԃ̂����鏈��
            fs.Close();

            Application.Run(new FormMain());
        }
    }
}