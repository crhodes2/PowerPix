**How to use the PowerPix Application**

PowerPix is a small application that allow user to use the power of the Google Search Engine to search for images on certain topics and use said image to create a PowerPoint Presentation slide with that image.

At this time, this version of PowerPix is for **developers only** and would require the following:
 -Visual Studio 2019 (to run and test the program)
 -Microsoft Office 365 (most preferably MS PowerPoint)
 -Access to a Google Custom Search API key.

To get started:

-Download the project on GitHub -\&gt; [https://github.com/crhodes2/PowerPix](https://github.com/crhodes2/PowerPix)

-Save and unzip the project somewhere on your Desktop

Get the API Key:

-In order to use the application, you need access to the Google Custom Search API Key. Here is how to do it:

1) Click on this link -\&gt; [https://developers.google.com/custom-search/v1/overview](https://developers.google.com/custom-search/v1/overview)

2) Scroll down until you see the **Get a Key** option ![](RackMultipart20200420-4-cpchpr_html_23ea16c9a4d13b0e.png)

3) You will be greeted with this pop up screen. Go ahead and click on **Create a new project** andname it whatever you&#39;d like. I&#39;d recommend naming it by the project name itself, which is in this case, PowerPix.

If you have a Google Account, proceed with signing in. Otherwise, creating a new account is free [https://support.google.com/accounts/answer/27441?hl=en](https://support.google.com/accounts/answer/27441?hl=en)

![](RackMultipart20200420-4-cpchpr_html_8a52606929362021.png)

4) After following these steps, you will be greeted with a box that has your API key on it. Copy that key on a clipboard.
 ![](RackMultipart20200420-4-cpchpr_html_7c33ed269cc3a736.png)

5) Go back to the project you have saved on your computer and navigate through to get to the file **App.Config.** Example: if the project was saved on your Desktop, it would be located under **&quot;C:\Users\username\Desktop\PowerPix-master\PowerPix2.0\App.config&quot;.** Open the file using Notepad.

6) Paste your key from your clipboard to the highlighted section. Save the file and close the Notepad.

![](RackMultipart20200420-4-cpchpr_html_343a00d490db14dd.png)
