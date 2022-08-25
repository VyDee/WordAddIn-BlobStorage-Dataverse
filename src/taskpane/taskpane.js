/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("searchButton").onclick = run;
    document.getElementById("allBtn").onclick = addAll;
  }
});

export async function run() {
  return Word.run(async (context) => {
    document.getElementById("loadingSpinner").classList.add("showAfterPicLoad");
    var componentOrderID = "" + document.getElementById("componentOrderIDInput").value;
    
    // generate the photo gallery
    var photoUrls = await getPhotosUrl(componentOrderID);

    const outElem = document.querySelector(".gallery");
    outElem.innerHTML = "";

    var itemsProcessed = 0;

    photoUrls.forEach(function (path) {
      const imgDiv = document.createElement("div");
      imgDiv.classList.add(
        "col-lg-4",
        "col-md-12",
        "mb-4",
        "mb-lg-0",
        "bg-image",
        "hover-overlay",
        "ripple",
        "shadow-1-strong",
        "rounded"
      );
      const img = document.createElement("img");
      img.src = path;
      img.classList.add("w-100", "h-100");
      img.onclick = function () {
        insertImage(this.src);
      };
      imgDiv.appendChild(img);
      outElem.appendChild(imgDiv);
      itemsProcessed++;
      if (itemsProcessed === photoUrls.length) {
        document.getElementById("loadingSpinner").classList.remove("showAfterPicLoad");
        //document.getElementById("allBtn").classList.add("showAfterPicLoad");
      }
    });

    await context.sync();
  });
}

export async function addAll() {
  return Word.run(async (context) => {
    var componentOrderID = "" + document.getElementById("componentOrderIDInput").value;
    var photoUrls = await getPhotosUrl(componentOrderID);

    for (let i = 0; i < photoUrls.length; i++) {
      await insertImage(photoUrls[i])    
    };

    await context.sync();
  });
}

function toDataURL(src, callback, outputFormat) {
  let image = new Image();
  image.crossOrigin = "Anonymous";
  image.onload = function () {
    console.log("Onload::toDataURL")
    let canvas = document.createElement("canvas");
    let ctx = canvas.getContext("2d");
    let dataURL;
    canvas.height = this.naturalHeight;
    canvas.width = this.naturalWidth;
    ctx.drawImage(this, 0, 0);
    dataURL = canvas.toDataURL(outputFormat);
    callback(dataURL);
  };
  image.src = src;
}

export async function insertImage(src) {
  await Word.run(async (context) => {
    toDataURL(src, function (dataUrl) {
      const base64Image = dataUrl.split(",")[1];
      context.document.body.paragraphs
        .getLast()
        .insertParagraph("", "After")
        .insertInlinePictureFromBase64(base64Image, "End");
      
      context.sync();
    });
  });
}

export async function getPhotosUrl(compOrID) {
  var request = new XMLHttpRequest();

  var apiUrl = process.env.RENEW_FA_API_URL;
  var param = "componentOrderID=" + compOrID;
  var functionKey = process.env.RENEW_FA_KEY;
  var url = apiUrl + "?" + functionKey + "&" + param;

  request.open("GET", url, false);
  request.setRequestHeader("Accept", "application/json");
  request.setRequestHeader("X-Request", "JSON");
  request.setRequestHeader("Access-Control-Allow-Origin", "*");
  request.send(null);

  if (Object.keys(request.response).length) {
    return JSON.parse(request.response).imageUrls;
  }

  return null;
}