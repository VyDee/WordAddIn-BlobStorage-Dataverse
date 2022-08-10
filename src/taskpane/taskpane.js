/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
//import "dotenv/config";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    // insert a paragraph at the end of the document.
    var photoUrls = await getPhotosUrl("203");
    const outElem = document.querySelector(".gallery");

    photoUrls.forEach(function (path) {
      const img = document.createElement("img");
      img.src = path;
      img.width = 100;
      img.length = 100;
      img.onclick = function () {
        insertImage(this.src);
      };
      outElem.appendChild(img);
    });

    await context.sync();
  });
}

function toDataURL(src, callback, outputFormat) {
  let image = new Image();
  image.crossOrigin = "Anonymous";
  image.onload = function () {
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
  if (image.complete || image.complete === undefined) {
    image.src = src;
  }
}

export async function insertImage(src) {
  await Word.run(async (context) => {
    toDataURL(src, function (dataUrl) {
      const base64Image = dataUrl.split(",")[1];
      context.document.body.paragraphs
        .getLast()
        .insertParagraph("", "After")
        .insertInlinePictureFromBase64(base64Image, "End");
    });

    await context.sync();
  });
}

export async function getPhotosUrl(compOrID) {
  var request = new XMLHttpRequest();
  // const dotenv = require("dotenv");
  // dotenv.config();
  // console.log(`RENEW key is ${process.env.RENEW_FA_KEY}`);

  var apiUrl = "Your Azure API Url";
  var param = "componentOrderID=" + compOrID;
  var functionKey = "Your Azure Function Key";
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
