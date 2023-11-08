const express = require('express');
const puppeteer = require('puppeteer');
const updtprog = require("./update_files_xlsx");

const folderPath = __dirname + "\\Databases"; // Replace with the path to your folder
const serviceEndpoint = "http://localhost:3000/scrape"; // Replace with your service endpoint

const app = express();
const port = 3000;

app.use(express.json());

app.post('/scrape', async (req, res) => {
  const { url, cssSelectors } = req.body;

  if (!url || !cssSelectors) {
    return res.status(400).json({ error: 'Both URL and CSS name are required.' });
  }

  //console.log(url);

  try {
    const browser = await puppeteer.launch({headless: "new"});
    const page = await browser.newPage();

    // Navigate the page to a URL
    await page.goto(url);
    
    // Set screen size  
    await page.setViewport({width: 1080, height: 1024});

    let data = {};

    let cssSelectors2 = Object.entries(cssSelectors);
    for( var indx = 0; indx < cssSelectors2.length; indx++)
    {
      var keypair = cssSelectors2[indx];
      var cssName = keypair[1];
      try{
      await page.waitForSelector(cssName);
      keypair[1] = await page.$eval(cssName, (element) => element.innerText);
      }catch(error0){
        continue;
      }
      cssSelectors[keypair[0]] = keypair[1] ;
      //console.log(keypair);
    }

    await browser.close();
    if(cssSelectors2.length == 1)
    res.send(Object.entries(cssSelectors)[0][1]);
    else
    res.json({ ...cssSelectors });
  } catch (error) {
    res.status(500).json({ error: 'An error occurred while scraping the website.' });
  }
});

//const intervalInMilliseconds = 1000; // 1 second in milliseconds
//const intervalInMilliseconds = 60000; // 1 minute in milliseconds
//const intervalInMilliseconds = 600000; // 10 minutes in milliseconds
const intervalInMilliseconds = 3600000; // 1 hour in milliseconds
setInterval( async()=>{
  const now = new Date();
  const timestamp = now.toISOString(); // Get a timestamp in ISO 8601 format
  console.log(`${timestamp} update started.`);
  updtprog.readExcelFilesInFolderAndSendRequests(folderPath, serviceEndpoint);
}, intervalInMilliseconds )

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
