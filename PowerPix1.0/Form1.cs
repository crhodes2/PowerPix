using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using System.Json;
using System.Configuration;
using System.Diagnostics;
using Microsoft.Office.Interop.PowerPoint;
using System.Net;
using Microsoft.Office.Core;

namespace PowerPix1._0
{

    public partial class Form1 : Form
    {
        #region initcomponent
        public Form1()
        {
            InitializeComponent();
        }
        #endregion

        #region folder access
        // ======== CREATE A FOLDER FOR THE PPT AND ITS IMAGES ======== //

        public void FolderAccess()
        {
            bool exists = Directory.Exists(@"C:\PPT");
            bool exists2 = Directory.Exists(@"C:\PPT\images");


            if (!exists)
                Directory.CreateDirectory(@"C:\PPT");
            if(!exists2)
                Directory.CreateDirectory(@"C:\PPT\images");

        }
        #endregion

        WebClient webClient = new WebClient();
        HttpClient client = new HttpClient();
        public string[] pptPics = { "", "", "", "", "", "", "", "", "", "" };
        List<String> generatedFilenameArray = new List<String>() { @"C:\PPT\images\picture1.png", @"C:\PPT\images\picture2.png", @"C:\PPT\images\picture3.png",
                @"C:\PPT\images\picture4.png", @"C:\PPT\images\picture5.png", @"C:\PPT\images\picture6.png", @"C:\PPT\images\picture7.png", @"C:\PPT\images\picture8.png",
            @"C:\PPT\images\picture9.png", @"C:\PPT\images\picture10.png"};


        class GoogleImage
        {
            public String linkOfImage { get; set; }
            public String thumbnailLink { get; set; }
        }

        public static String userInput;         // the user input. Whatever the user types
        JsonValue pptJSON;                      // the JSON string taken from the Google API
        public String selectedPic;              // the picture selected by the user to add to their PPT slide
        GoogleImage image = new GoogleImage();  // image selections taken from the Google API


        #region unused methods
        private void label1_Click(object sender, EventArgs e)
        {
            // not used
        }
        #endregion

        private void searchBtn_Click(object sender, EventArgs e)
        {
            Request();
        }

        public async void Request()
        {
            userInput = searchBar.Text;
            if (userInput != "")
            {
                descriptionBox.Text = "...searching...";
                string API_KEY = ConfigurationManager.AppSettings["API_KEY"];
                string CUST_SEARCH_ID = ConfigurationManager.AppSettings["CUST_SEARCH_ID"];
                string url = string.Format("https://www.googleapis.com/customsearch/v1?key={0}&cx={1}&q={2}&searchType=image", API_KEY, CUST_SEARCH_ID, userInput);
                var myWebRequest = await client.GetStringAsync(url);
                string req = myWebRequest;
                results(req);
                descriptionBox.Text = userInput;

            }

            else
            {
                descriptionBox.Text = "404 - Not Found";
            }
        }

        public void results(string jsonData)
        {
            List<String> listOfImages = new List<String>(); //define a list of images
            pptJSON = JsonArray.Parse(jsonData);   //parse the json data as a result of the user input

            // IF ELSE STATEMENT
            // if we already have a list of images, clear it first. Otherwise, display your list of images based on the json data found
            if (listOfImages.Contains(image.linkOfImage))
            {
                listOfImages.Clear();
            }

            else
            {
                // take an image link found in your Google Search and add them all to your 10-item list of images. Stop the loop once the 10th item has been added.
                for (int i = 0; i < 10; i++)
                {
                    image.linkOfImage = pptJSON["items"][i]["link"];
                    Trace.WriteLine(image.linkOfImage);

                    listOfImages.Add(image.linkOfImage);
                }
            }

            FolderAccess();

            for (int a = 0; a < generatedFilenameArray.Count; a++)
            {
                webClient.DownloadFile(listOfImages[a], generatedFilenameArray[a]);
            }
            string generatedFileName = generatedFilenameArray[0];
            string generatedFileName2 = generatedFilenameArray[1];
            string generatedFileName3 = generatedFilenameArray[2];
            string generatedFileName4 = generatedFilenameArray[3];
            string generatedFileName5 = generatedFilenameArray[4];
            string generatedFileName6 = generatedFilenameArray[5];
            string generatedFileName7 = generatedFilenameArray[6];
            string generatedFileName8 = generatedFilenameArray[7];
            string generatedFileName9 = generatedFilenameArray[8];
            string generatedFileName10 = generatedFilenameArray[9];

            pictureBox1.ImageLocation = generatedFileName;
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox2.ImageLocation = generatedFileName2;
            pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox3.ImageLocation = generatedFileName3;
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox4.ImageLocation = generatedFileName4;
            pictureBox4.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox5.ImageLocation = generatedFileName5;
            pictureBox5.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox6.ImageLocation = generatedFileName6;
            pictureBox6.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox7.ImageLocation = generatedFileName7;
            pictureBox7.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox8.ImageLocation = generatedFileName8;
            pictureBox8.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox9.ImageLocation = generatedFileName9;
            pictureBox9.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox10.ImageLocation = generatedFileName10;
            pictureBox10.SizeMode = PictureBoxSizeMode.StretchImage;

        }

        private void create_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.Control c in this.Controls)
            {
                if ((c is System.Windows.Forms.CheckBox))
                {
                    if (((System.Windows.Forms.CheckBox)c).Checked)
                    {
                        pptPics[int.Parse(c.Name[8].ToString()) - 1] = generatedFilenameArray[int.Parse(c.Name[8].ToString()) - 1];
                    }
                    else
                    {
                        pptPics[int.Parse(c.Name[8].ToString()) - 1] = "";
                    }
                }
            }

            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            CustomLayout custLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

            var slides = pptPresentation.Slides;
            _Slide slide = slides.AddSlide(1, custLayout);
            
            //Create Title 
            var objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = searchBar.Text;
            objText.Font.Size = 32;
            slide.Shapes[1].Height = 60;
            
            //Create Description
            var objText2 = slide.Shapes[2].TextFrame.TextRange;
            objText2.Text = descriptionBox.Text;
            objText2.Font.Size = 16;
            slide.Shapes[2].Width = 310;
            slide.Shapes[2].Top = 115;

            //Add Images to the Slides
            int height = 200;
            int width = 155;
            int verticalPosition = 115;
            int horizontalPosition = 370;
            int position = 1;

            for (int i = 0; i < pptPics.Length; i++)
            {
                if (pptPics[i] == "")
                    continue;

                verticalPosition = (position == 1 || position == 2) ? 115 : 315;
                horizontalPosition = (position == 1 || position == 3) ? 370 : 525;


                slide.Shapes.AddPicture(
                    pptPics[i],
                    MsoTriState.msoFalse,
                    MsoTriState.msoTrue,
                    horizontalPosition,
                    verticalPosition,
                    width,
                    height);

                verticalPosition += height + 5;
                position++;
            }

            FolderAccess();

            int slideNumber = 1;
            string filePath = @"C:\PPT\newslide1.pptx";

            while (File.Exists(filePath))
            {
                slideNumber += 1;
                filePath = @"C:\PPT\newslide" +
                           slideNumber.ToString() +
                           ".pptx";
            }
            pptPresentation.SaveAs(
                filePath,
                PpSaveAsFileType.ppSaveAsDefault,
                MsoTriState.msoTrue);
        }
    }
}

