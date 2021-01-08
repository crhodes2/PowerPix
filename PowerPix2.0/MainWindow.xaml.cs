using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Net.Http;
using System.Json;
using System.Configuration;
using System.Diagnostics;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

/*
 *  PowerPix
     A program that utilizes Google Custom Search API
     to search for user-specified images to create and add 
     into a PowerPoint presentation.

*/

namespace PowerPix2._0
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public static String userInput;         // the user input. Whatever the user types
        JsonValue pptJSON;                      // the JSON string taken from the Google API
        public String selectedPic;              // the picture selected by the user to add to their PPT slide
        GoogleImage image = new GoogleImage();  // image selections taken from the Google API

        public MainWindow()
        {
            InitializeComponent();
        }


        private void searchBtn_Click(object sender, RoutedEventArgs e)
        {
            Request();
        }

        public async void Request()
        {
            userInput = searchBox.Text;
            if (userInput != "")
            {
                searchResults.Content = "...searching...";
                string API_KEY = ConfigurationManager.AppSettings["API_KEY"];
                string CUST_SEARCH_ID = ConfigurationManager.AppSettings["CUST_SEARCH_ID"];
                var client = new HttpClient();
                string url = string.Format("https://www.googleapis.com/customsearch/v1?key={0}&cx={1}&q={2}&searchType=image", API_KEY, CUST_SEARCH_ID, userInput);
                var myWebRequest = await client.GetStringAsync(url);
                string req = myWebRequest;
                results(req);
                searchResults.Content = "Search Results for: " + userInput;

            }

            else
            {
                searchResults.Content = "404 - Not Found";
            }
        }


        public void results(string jsonData)
        {
            pptJSON = JsonArray.Parse(jsonData);
            List<String> listOfImages = new List<String>();

            if (listOfImages.Contains(image.linkOfImage))
            {
                listOfImages.Clear();
            }

            else
            {
                for (int i = 0; i < 10; i++)
                {
                    image.linkOfImage = pptJSON["items"][i]["link"];
                    Trace.WriteLine(image.linkOfImage);

                    listOfImages.Add(image.linkOfImage);
                }

                googleListView.ItemsSource = listOfImages;
            }
        }


        private void googleListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            //https://docs.microsoft.com/en-us/dotnet/api/system.windows.controls.image.source?view=netframework-4.8#System_Windows_Controls_Image_Source
            int i = e.AddedItems.Count;

            BitmapImage bi3 = new BitmapImage();

            bi3.BeginInit();

            bi3.UriSource = new Uri(e.AddedItems[0].ToString());
            selectedPic = e.AddedItems[0].ToString();

            bi3.EndInit();

            img.Stretch = Stretch.Fill;
            img.Source = bi3;

        }

        private void googleListView_Scroll(object sender, System.Windows.Controls.Primitives.ScrollEventArgs e)
        {

        }

        private void createSlideBtn_Click(object sender, RoutedEventArgs e)
        {
            userInput = searchBox.Text;
            if (userInput != "")
            {
                string pictureFile = selectedPic;
                string userDir = DateTime.Now.Ticks.ToString();

                // Create the PowerPoint Presentation
                Microsoft.Office.Interop.PowerPoint.Application pptApplication =
                    new Microsoft.Office.Interop.PowerPoint.Application();
                Microsoft.Office.Interop.PowerPoint.Slides slides;
                Microsoft.Office.Interop.PowerPoint._Slide slide;
                Microsoft.Office.Interop.PowerPoint.TextRange objText;
                Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
                Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

                slides = pptPresentation.Slides;
                slide = slides.AddSlide(1, customLayout);

                Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];

                slide.Shapes.AddPicture(pictureFile, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);
                //pptPresentation.SaveAs(@"C:\\" + userDir  + ".pptx", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoTrue);
                //pptPresentation.Close();
                //pptApplication.Quit();
            }

            else
            {
                searchResults.Content = "Cannot create PPT. Please perform a search query";
            }

        }
    }
}
