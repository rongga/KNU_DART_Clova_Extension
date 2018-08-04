module.exports = function(app){
  const express = require('express');
  const router = express.Router();
  const xlsx = require('xlsx');
  const fs = require('fs');
  const bodyParser = require('body-parser');




  app.use(bodyParser.json());

  router.get('/',function(req,res){
    res.send('Hi');
  });
  router.post('/',function(req,res){

    const body = req.body;
    const type = body.request.type;
    var resStr = initRes();
    console.log(type+" is arrived.");
    switch (type) {
      case "EventRequest":
        res.send("Event");
        break;
      case "IntentRequest":

        var name = body.request.intent.name;

        if(name == "Init"){
          makeRes(resStr,"검색이 초기화 되었습니다. 다시 검색해 주세요.");
          res.send(resStr);

        }else if(name == "List"){



          var slots = body.request.intent.slots;
          var indicator = slots.Indicator.value;
          var workbook = xlsx.readFile("public/"+indicator+".xlsx");
          var worksheet = workbook.Sheets["Sheet1"];

          var standard = slots.Standard.value;
          var condition = slots.Condition.value;

          if(slots.Date){
            var date = slots.Date.value;
          }else{
            var date = 2017;
          }

          var conKey = readCon(condition);
          var dateChar = findDateChar(date,worksheet);

          var conList = [];


          var List = [];
          if(body.session.sessionAttributes){
            if(body.session.sessionAttributes.List){
              List = body.session.sessionAttributes.List;
            }
            if(body.session.sessionAttributes.conList){
              conList = body.session.sessionAttributes.conList;
            }
          }
          conList.push({
            "indicator" : indicator,
            "standard" : standard,
            "condition": condition
          });

          if(List.length > 0){
            console.log(conKey);
            var tmp = [];
            for(var i = 0; i < List.length; i++){
              for(var j = 2; j < 2040; j++){
                if(worksheet["A"+j].v == List[i]){
                  if(conKey == 1){
                    if(worksheet[dateChar+j] && worksheet[dateChar+j].v >= standard){
                      tmp.push(worksheet["A"+j].v);
                    }
                  }else if(conKey == 2){
                    if(worksheet[dateChar+j] && worksheet[dateChar+j].v <= standard){
                      tmp.push(worksheet["A"+j].v);
                    }
                  }else if(conKey == 3){
                    if(worksheet[dateChar+j] && worksheet[dateChar+j].v > standard){
                      tmp.push(worksheet["A"+j].v);
                    }
                  }else if(conKey == 4){
                    if(worksheet[dateChar+j] && worksheet[dateChar+j].v < standard){
                      tmp.push(worksheet["A"+j].v);
                    }
                  }
                }
              }

            }

            List = tmp;
          }else if (List.length == 0) {
            console.log(conKey);

            for(var j = 2; j < 2040; j++){

                if(conKey == 1){
                  if(worksheet[dateChar+j] && worksheet[dateChar+j].v >= standard){
                    List.push(worksheet["A"+j].v);
                  }
                }else if(conKey == 2){
                  if(worksheet[dateChar+j] && worksheet[dateChar+j].v <= standard){
                    List.push(worksheet["A"+j].v);
                  }
                }else if(conKey == 3){
                  if(worksheet[dateChar+j] && worksheet[dateChar+j].v > standard){
                    List.push(worksheet["A"+j].v);
                  }
                }else if(conKey == 4){
                  if(worksheet[dateChar+j] && worksheet[dateChar+j].v < standard){
                    List.push(worksheet["A"+j].v);
                  }
                }

            }
            console.log(List.length);
          }

          resStr.sessionAttributes = {
            "List": List,
            "conList": conList
          };
          if(List.length > 0){
            makeRes(resStr,List.length + "개의 종목이 검색되었습니다. 추가로 검색하시거나 처음으로 돌아가시려면 \"다시\"라고 말해주세요.");
          }else if(List.length == 0){
            makeRes(resStr,"검색된 종목이 없습니다. 다시 검색해 주세요.");
          }
          resStr.response.card = {
            "type": "Text",
            "mainText": {
              "type": "string",
              "value": List.join('\n')
            }
          };

          res.send(resStr);
          console.log(indicator + " " + standard + condition + " 주식 입니다.");
        }else if(name == "Search"){

          var slots = body.request.intent.slots;
          var indicator = slots.Indicator.value;
          var workbook = xlsx.readFile("public/"+indicator+".xlsx");
          var worksheet = workbook.Sheets["Sheet1"];

          var code = slots.Code.value;
          if(slots.Date){
            var date = slots.Date.value;
          }else{
            var date = 2017;
          }
          var dateChar = findDateChar(date,worksheet);
          var key = "자료없음";

          for(var i = 2; i < 2040; i++){
            if(worksheet[dateChar+i] && worksheet["A"+i].v == code){
              key = worksheet[dateChar+i].v;
              break;
            }
          }


          if(key == "자료없음"){
            makeRes(resStr,"찾으시는 자료가 없습니다. 다시 검색해 주세요.");
          }else{
            makeRes(resStr,date + "년 " + code + "의 " + indicator + "은 " + key + "입니다.");
          }
          res.send(resStr);
        }else if (name == "Rank") {

          var slots = body.request.intent.slots;
          var indicator = slots.Indicator.value;
          var workbook = xlsx.readFile("public/"+indicator+".xlsx");
          var worksheet = workbook.Sheets["Sheet1"];
          var rankCon = slots.RankCon.value;
          var num = slots.Num.value;
          if(slots.Date){
            var date = slots.Date.value;
          }else{
            var date = 2017;
          }
          var dateChar = findDateChar(date,worksheet);

          var conKey = readCon(rankCon);

          var rankList = [];
          var tmpN = 2040;
          var ti = 3;
          for(var j = 0; j<num; j++){
            tmp = worksheet[dateChar+(ti+1)].v || worksheet[dateChar+(ti-1)].v;
            for(var i = 2; i < tmpN; i++){

              if(conKey == 1){
                if(worksheet[dateChar+i] && worksheet[dateChar+i].v > tmp){
                  tmp = worksheet[dateChar+i].v;
                  ti = i;
                }
              }else if (conKey == 2) {
                if(worksheet[dateChar+i] && worksheet[dateChar+i].v < tmp){
                  tmp = worksheet[dateChar+i].v;
                  ti = i;
                }
              }
            }

            console.log(worksheet["A"+ti].v);
            if(worksheet["A"+ti]){
              rankList.push(worksheet["A"+ti].v);
            }
            delete worksheet[dateChar+ti];
            tmpN--;
          }
          makeRes(resStr,rankList.join(', ')+"가 검색되었습니다.");
          res.send(resStr);
        }else if (name == "BackTest") {
          if(body.session.sessionAttributes){
            if(body.session.sessionAttributes.conList){

              var tmpList = [];
              var conList = body.session.sessionAttributes.conList;
              var worksheets = [];
              for(var i = 0; i < conList.length; i++){
                worksheets[i] = xlsx.readFile("public/"+conList[i].indicator+".xlsx").Sheets["Sheet1"];
              }
              var returnSheet = xlsx.readFile("public/return.xlsx").Sheets["Sheet1"];

              var cum = [1];
              for(var i = 1; i < 7; i++){
                var tmpDate = String.fromCharCode(i+65);
                var tmpR = 0;
                for(var j = 0; j < tmpList.length; j++){
                   //sell & calculate rate
                   for(var CDn = 2; CDn < 2040; CDn++){
                     if(returnSheet[tmpDate+CDn] && returnSheet['A'+CDn].v == tmpList[j]){
                       tmpR += returnSheet[tmpDate+CDn].v/100;
                     }
                   }
                }
                if(tmpList.length > 0){
                  tmpR /= tmpList.length;
                }



                cum.push(cum[cum.length-1]*(1+tmpR));
                tmpList.length = 0;

                for(var CLn = 0; CLn < conList.length; CLn++){
                  var tmpConKey = readCon(conList[CLn].condition);
                  var tmp2 = [];
                  for(var CDn = 2; CDn < 2040; CDn++){
                    if(tmpConKey == 1){
                      if(CLn == 0){
                        if(worksheets[CLn][tmpDate+CDn] && worksheets[CLn][tmpDate+CDn].v >= conList[CLn].standard){
                          tmpList.push(worksheets[CLn]['A'+CDn].v);
                        }
                      }else{
                        for(var tLn = 0; tLn < tmpList.length; tLn++){
                          if(worksheets[CLn][tmpDate+CDn] && (worksheets[CLn]['A'+CDn].v == tmpList[tLn]) && (worksheets[CLn][tmpDate+CDn].v >= conList[CLn].standard)){
                            tmp2.push(worksheets[CLn]['A'+CDn].v);
                          }
                        }
                      }
                    }else if (tmpConKey == 2) {
                      if(CLn == 0){
                        if(worksheets[CLn][tmpDate+CDn] && worksheets[CLn][tmpDate+CDn].v <= conList[CLn].standard){
                          tmpList.push(worksheets[CLn]['A'+CDn].v);
                        }
                      }else{
                        for(var tLn = 0; tLn < tmpList.length; tLn++){
                          if(worksheets[CLn][tmpDate+CDn] && (worksheets[CLn]['A'+CDn].v == tmpList[tLn]) && (worksheets[CLn][tmpDate+CDn].v <= conList[CLn].standard)){
                            tmp2.push(worksheets[CLn]['A'+CDn].v);
                          }
                        }
                      }
                    }else if (tmpConKey == 3) {
                      if(CLn == 0){
                        if(worksheets[CLn][tmpDate+CDn] && worksheets[CLn][tmpDate+CDn].v > conList[CLn].standard){
                          tmpList.push(worksheets[CLn]['A'+CDn].v);
                        }
                      }else{
                        for(var tLn = 0; tLn < tmpList.length; tLn++){
                          if(worksheets[CLn][tmpDate+CDn] && (worksheets[CLn]['A'+CDn].v == tmpList[tLn]) && (worksheets[CLn][tmpDate+CDn].v > conList[CLn].standard)){
                            tmp2.push(worksheets[CLn]['A'+CDn].v);
                          }
                        }
                      }
                    }else if (tmpConKey == 4) {
                      if(CLn == 0){
                        if(worksheets[CLn][tmpDate+CDn] && worksheets[CLn][tmpDate+CDn].v < conList[CLn].standard){
                          tmpList.push(worksheets[CLn]['A'+CDn].v);
                        }
                      }else{
                        for(var tLn = 0; tLn < tmpList.length; tLn++){
                          if(worksheets[CLn][tmpDate+CDn] && (worksheets[CLn]['A'+CDn].v == tmpList[tLn]) && (worksheets[CLn][tmpDate+CDn].v < conList[CLn].standard)){
                            tmp2.push(worksheets[CLn]['A'+CDn].v);
                          }
                        }
                      }
                    }
                  }
                  if(CLn != 0){
                    tmpList = tmp2;
                  }
                  // console.log(tmpList+", "+tmpDate);
                }
              }
              var e = 1;
              var MDD=0;
              var et = [];
              for(var i = 1;i < cum.length;i++){
                if(cum[i] < cum[i-1]){
                  e = e*(1 + ((cum[i]-cum[i-1])/cum[i-1]));
                }else{
                  if((1-e) > MDD){
                    MDD = 1-e;
                  }
                  e=1;
                }

                et.push(e);
              }

              var avg = 100*(Math.pow(cum[cum.length-1],1/5)-1);
              MDD *= 100;
              makeRes(resStr,"백테스팅 결과, 평균복리수익률은 "+avg.toFixed(2)+"퍼센트, MDD는 "+MDD.toFixed(2)+"퍼센트 입니다.");
            }
          }else {
            makeRes(resStr,"저장된 조건이 없습니다. 처음부터 검색해 주세요.");
          }
          res.send(resStr);

        }else{
          makeRes(resStr,"잘 알아듣지 못했습니다. 다시 말씀해주세요.");
          res.send(resStr);
        }


        break;

      case "LaunchRequest":
        makeRes(resStr,"다트 종목검색기를 시작합니다.");
        res.send(resStr);
        break;

      case "SessionEndedRequest":
        makeRes(resStr,"다트 종목검색기를 종료합니다.");
        resStr.shouldEndSession = true;
        res.send(resStr);
        break;

      default:

        break;
    }

  });
  return router;
};
function initRes(){
  var resStr =
  {
    "version": "0.0.1",
    "sessionAttributes": {},
    "response": {
      "outputSpeech": {
        "type": "SimpleSpeech",
        "values": {
            "type": "PlainText",
            "lang": "ko",
            "value": ""
        }
      },
      "card": {},
      "directives": [],
      "shouldEndSession": false
    }
  };
  return resStr;
}
function readCon(condition){
  if(condition == "이상" || condition == "상위"){
    return 1;
  }else if (condition == "이하" || condition == "하위") {
    return 2;
  }else if (condition == "초과") {
    return 3;
  }else if (condition == "미만"){
    return 4;
  }else{
    return -1;
  }
}
function findDateChar(date,worksheet) {
  var dateChar;
  for(var i = 1; i < 7; i++){
    dateChar = String.fromCharCode(i+65);
    if(worksheet[dateChar+"1"].v == date){
      break;
    }
  }
  return dateChar;
}

function makeRes(resStr,string){
  resStr.response.outputSpeech.values = {
    "type": "PlainText",
    "lang": "ko",
    "value": string
  };
}
