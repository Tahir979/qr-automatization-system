using System;
using System.Threading.Tasks;

namespace QR_Automatization
{
    public partial class loadingscreen : MetroFramework.Forms.MetroForm
    {
        public Action Worker { get; set; }
        public loadingscreen(Action worker)
        {
            InitializeComponent();
            if (worker == null)
            {
                throw new ArgumentNullException();
            }
            Worker = worker;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Task.Factory.StartNew(Worker).ContinueWith(t => { this.Close(); }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void loadingscreen_Load(object sender, EventArgs e)
        {
            if (label.Default.text == "Loading...")
            {
                text.Text = "Loading...";
            }
            else if (label.Default.text == "Creating...")
            {
                text.Text = "Creating...";
            }
            else if (label.Default.text == "Merging and Saving...")
            {
                text.Text = "Merging and Saving...";
            }
            else if (label.Default.text == "Uploading...")
            {
                text.Text = "Uploading...";
            }
            else if (label.Default.text == "Loading Camera...")
            {
                text.Text = "Loading Camera...";
            }
        }
    }
}
