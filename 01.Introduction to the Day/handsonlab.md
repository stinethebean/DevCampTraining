##Create an Excel add-in using the Napa Development Tools
### Go to the Napa Tools
1. Navigate to [https://www.napacloudapp.com/Getting-Started/](https://www.napacloudapp.com/Getting-Started)
2. Log in ![](http://i.imgur.com/S3g6syA.jpg)

### Create an add-in for Excel
1. Select **Add New Project** ![](http://i.imgur.com/wcbfxto.png)
2.
3. Click on the Task Pane app for Office option. Name your Task Pane Project and click create.![](http://i.imgur.com/7PeKxvs.png)
4. Delete Everything inside the `<body></body>` tag
![](http://i.imgur.com/0Syg3sD.png)

5. Add the following code
 
	`<div id="content-header"><div class="padding"><h1>Welcome!</h1>        </div></div><div id="content-main">      <div class="padding">          <p><strong>Select text and find related Flickr images.</strong></p>                   <button id="get-data-from-selection">Search Flickr</button>      </div><div id="Images"></div></div>`

6. Navigate to Home.js.    In `if (result.status === Office.AsyncResultStatus.Succeeded) {` replace the content with : 
`app.showNotification('The selected text is:', '"' + result.value + '"');                showImages(result.value);`
![](http://i.imgur.com/I1FkZeW.png)

7. Add the showImages Function
    `function showImages(selectedText) {$('#Images').empty();var parameters = {tags: selectedText,tagsmode: "any",format: "json"};$.getJSON("https://secure.flickr.com/services/feeds/photos_public.gne?jsoncallback=?", parameters,function (results) {$.each(results.items, function (index, item) {$('#Images').append($("<img style='height:100px; width: auto; padding-right: 5px;'/>").attr("src", item.media.m));});});}`
![](http://i.imgur.com/bSa61w7.png)

8. Run the add in.

![](http://i.imgur.com/05iRkXI.png)

9. Start the App, and test it out!

![](http://i.imgur.com/Klmu40F.png)
![](http://i.imgur.com/9nnTsJJ.png)
