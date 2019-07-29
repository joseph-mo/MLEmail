"use strict";

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

var h = require("http");
var accessKey = "4d93eb8eb78d4062966bb0a4c96ff7d7";

var uri = "eastus.api.cognitive.microsoft.com";
var generalPath = "/text/analytics/v2.1/";
var sentimentPath = "/text/analytics/v2.1/sentiment";
var langPath = "/text/analytics/v2.1/languages";
var keyPhrasePath = "/text/analytics/v2.1/keyPhrases";

var retVal;
var scores = [];
var langScore = [];
var phrases = [];
var language = "English";

let response_handler = function(response) {
  let body = "";
  response.on("data", function(d) {
    body += d;
  });

  response.on("end", function() {
    let body_ = JSON.parse(body);

    console.log(body_);
    if ("score" in body_["documents"][0]) {
      scores.push(body_["documents"][0]["score"]);
    } else if ("keyPhrases" in body_["documents"][0]) {
      phrases.push(body_["documents"][0]["keyPhrases"].join(", ").split(","));
    } else if ("detectedLanguages" in body_["documents"][0]) {
      language = body_["documents"][0]["detectedLanguages"][0]["name"];
      langScore.push(body_["documents"][0]["detectedLanguages"][0]["score"]);
    }
    var body__ = JSON.stringify(body_, null, "  ");

    // retVal = body__;
  });
  response.on("error", function(e) {
    console.log("Error: ", e.message);
  });
};

// Actual sentiment of the body of the email
let get_sentiments = function(documents) {
  let body = JSON.stringify(documents);
  let request_params = {
    method: "POST",
    hostname: uri,
    path: sentimentPath,
    headers: {
      "Content-Type": "application/json",
      "Ocp-Apim-Subscription-Key": accessKey
    },
    body: undefined
  };

  let req = h.request(request_params, response_handler);
  req.write(body);
  req.end();
};

let get_key_phrases = function(documents) {
  let body = JSON.stringify(documents);

  let request_params = {
    method: "POST",
    hostname: uri,
    path: keyPhrasePath,
    headers: {
      "Content-Type": "application/json",
      "Ocp-Apim-Subscription-Key": "4d93eb8eb78d4062966bb0a4c96ff7d7"
    },
    body: undefined
  };

  let req = h.request(request_params, response_handler);
  req.write(body);
  req.end();
};

let detect_language = function(documents) {
  let body = JSON.stringify(documents);

  let request_params = {
    method: "POST",
    hostname: uri,
    path: langPath,
    headers: {
      "Content-Type": "application/json",
      "Ocp-Apim-Subscription-Key": accessKey
    },
    body: undefined
  };
  let req = h.request(request_params, response_handler);
  req.write(body);
  req.end();
};

var counter = 1;
export async function run() {
  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
      var documents = {
        documents: [
          {
            id: counter.toString(),
            // This retrieves body of email
            text: result.value
          }
        ]
      };
      counter++;
      get_sentiments(documents);
      get_key_phrases(documents);
      detect_language(documents);
      document.getElementById("btn").onclick = function() {
        if (language === "English") {
          document.getElementById("item-subject").innerHTML =
            "<b>Sentiment Rating:</b> <br/>" +
            Math.round(1000 * scores[scores.length - 1]) / 1000;
        } else {
          document.getElementById("item-subject").innerHTML =
            "<b>Sentiment Rating:</b> <br/>" +
            Math.round(1000 * langScore[langScore.length - 1]) / 1000;
        }

        // Displays emoji based off where score falls in range
        var img = document.getElementById("emoji1");
        var img1 = document.getElementById("emoji2");
        var img2 = document.getElementById("emoji3");
        console.log(scores);
        console.log(langScore);
        if (
          (language == "English" && scores[scores.length - 1] > 0.7) ||
          (language != "English" && langScore[langScore.length - 1] > 0.7)
        ) {
          img.style.display = "block";
        } else if (
          (language == "English" &&
            scores[scores.length - 1] > 0.3 &&
            scores[scores.length - 1] < 0.7) ||
          (language != "English" &&
            langScore[langScore.length - 1] > 0.3 &&
            langScore[langScore.length - 1] < 0.7)
        ) {
          img2.style.display = "block";
        } else {
          img1.style.display = "block";
        }
        document.getElementById("item-footer").innerHTML =
          "<b>Key Phrases: </b> <br/> <p>" +
          phrases[phrases.length - 1] +
          "</p>";
      };
    }
  );
}

// Office.context.mailbox.addHandlerAsync(
//   Office.EventType.ItemChanged,
//   function itemChanged(eventArgs) {
//     UpdateTaskPaneUI(Office.context.mailbox.item);
//   },
//   function UpdateTaskPaneUI(item) {
//     // Assuming that item is always a read item (instead of a compose item).
//     if (item != null) console.log(item.subject);
//   }
// );
