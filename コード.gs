function getProperty(key){
  let token = PropertiesService.getScriptProperties().getProperty(key);
  return token;
}

let form = FormApp.openById(getProperty('form1'));
let form2 = FormApp.openById(getProperty('form2'));
let form3 = FormApp.openById(getProperty('form3'));
let form4 = FormApp.openById(getProperty('form4'));
let ss = SpreadsheetApp.openById(getProperty('ss'));

function init_form(){
  //初期化
  deleteItems(form);
  deleteItems(form2);
  deleteItems(form3);
  deleteItems(form4);
}

function makeXmasQuestion() {
  var sheet = ss.getSheetByName('データ入力');
  var lastRow1 = sheet.getRange('C69').getValue();
  var lastRow2 = sheet.getRange('I140').getValue();
  //二次元配列
  var nameIndex1 = sheet.getRange(5,2,lastRow1,1).getValues();
  var nameIndex2 = sheet.getRange(5,3,lastRow1,1).getValues();
  var nameIndex3 = sheet.getRange(5,6,lastRow1,1).getValues();
  var questionIndex = sheet.getRange(5,8,lastRow2,1).getValues();

  //全員リスト
  var namesAll = [];
  for (var i in nameIndex1){
    namesAll.push(nameIndex1[i][0])
  };
  //男子リスト
  var namesMen = [];
  for (var i in nameIndex2){
    if(nameIndex2[i][0]!=''){
    namesMen.push(nameIndex2[i][0])
    }
  };
  //女子リスト
  var namesWomen = [];
  for (var i in nameIndex3){
    if(nameIndex3[i][0]!=''){
      namesWomen.push(nameIndex3[i][0])
    }
  };

  //Logger.log(namesMen)
  //フォーム作成
  form.addListItem().setTitle('名前を選択').setChoiceValues(namesAll).setRequired(true);
  form2.addListItem().setTitle('名前を選択').setChoiceValues(namesAll).setRequired(true);
  form3.addListItem().setTitle('名前を選択').setChoiceValues(namesWomen).setRequired(true);
  form4.addListItem().setTitle('名前を選択').setChoiceValues(namesMen).setRequired(true);

  var no = 0;//質問全体の通しNo
  var no2 = 0;//フォーム1の何ページ目か？
  var no3 = 0;//フォーム２の何ページ目か？
  for(var i in questionIndex){
    var question = questionIndex[i][0];
    no += 1;
    //if(no <= 63){//フォーム１
    if(no <= lastRow2/2){//フォーム１は半分の質問
      
      if(question == '理想のMIXペア'){//「理想のMIXペア」だけアンケートの形が違うので別処理
        form.addListItem().setTitle('No.'+no+'\n'+question+'(1ペア目1人目)').setChoiceValues(namesAll).setRequired(true);
        form.addListItem().setTitle(question+'(1ペア目2人目)').setChoiceValues(namesAll).setRequired(true);
        form.addListItem().setTitle(question+'(2ペア目1人目)').setChoiceValues(namesAll).setRequired(true);
        form.addListItem().setTitle(question+'(2ペア目2人目)').setChoiceValues(namesAll).setRequired(true);
      //}else if(no % 25 == 1){//質問25個ごとにページを分ける
        }else if(no % 10 == 1){//(大体)質問10個ごとにページを分ける
          
        no2 += 1;
        form.addPageBreakItem().setTitle('アンケートページ'+'('+no2+'ページ目)');
        form.addListItem().setTitle('No.'+no+'\n'+question+'(1人目)').setChoiceValues(namesAll).setRequired(true);
        form.addListItem().setTitle(question+'(2人目)').setChoiceValues(namesAll).setRequired(true);
      }else{
        form.addListItem().setTitle('No.'+no+'\n'+question+'(1人目)').setChoiceValues(namesAll).setRequired(true);
        form.addListItem().setTitle(question+'(2人目)').setChoiceValues(namesAll).setRequired(true);
      }
    }else{//フォーム２
      //if(no % 25 ==14){//質問64
      if(no % 10 ==1){//（だいたい）質問10個ごとにページを分ける
        
      no3 += 1;
      form2.addPageBreakItem().setTitle('アンケートページ'+'('+no3+'ページ目)');
      form2.addListItem().setTitle('No.'+no+'\n'+question+'(1人目)').setChoiceValues(namesAll).setRequired(true);
      form2.addListItem().setTitle(question+'(2人目)').setChoiceValues(namesAll).setRequired(true);
    }else{
      form2.addListItem().setTitle('No.'+no+'\n'+question+'(1人目)').setChoiceValues(namesAll).setRequired(true);
      form2.addListItem().setTitle(question+'(2人目)').setChoiceValues(namesAll).setRequired(true);
    }
    }
  }
  
  
  //推し男・シケ男
  form3.addPageBreakItem().setTitle('推し男');
  for(var i = 1;i < 11;i++){
    form3.addListItem().setTitle(i +'位').setChoiceValues(namesMen).setRequired(true);
    form3.addParagraphTextItem().setTitle('理由・コメント').setRequired(true);
  }
  form3.addPageBreakItem().setTitle('シケ男');
  for(var i = 1;i < 4;i++){
    form3.addListItem().setTitle(i +'位').setChoiceValues(namesMen).setRequired(true);
    form3.addParagraphTextItem().setTitle('理由・コメント').setRequired(true);
  }
  //男子が選ぶ推し男・シケ男
  form4.addPageBreakItem().setTitle('男子が選ぶ推し男');
  for(var i = 1;i < 11;i++){
    form4.addListItem().setTitle(i +'位').setChoiceValues(namesMen).setRequired(true);
    form4.addParagraphTextItem().setTitle('理由・コメント').setRequired(true);
  }
  form4.addPageBreakItem().setTitle('男子が選ぶシケ男');
  for(var i = 1;i < 4;i++){
    form4.addListItem().setTitle(i +'位').setChoiceValues(namesMen).setRequired(true);
    form4.addParagraphTextItem().setTitle('理由・コメント').setRequired(true);
  }

}

function deleteItems(a){
  //var a = form1;
  var items = a.getItems();
  var end = items.length - 1;
  for(var i = end ; i >= 0; i--){
    a.deleteItem(i);
  }
}

//計算タブ専用
function myFunction(){
  var box = [];
  var alpha = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z'];
  var sheet = ss.getSheetByName('計算');
  
  var sheet3 = ss.getSheetByName('データ入力');//
  var lastRow1 = sheet3.getRange('C69').getValue();//
  var lastRow2 = sheet3.getRange('I140').getValue();//
  
  for (var i in alpha){
    box.push(alpha[i]);
  }
  for (var i in alpha){
    for (var j in alpha){
      box.push(alpha[i]+alpha[j]);
      }
  }
  //for (var i = 1 ; i<=126 ; i++){
  for (var i = 1 ; i<=lastRow2 ; i++){//全ての質問を回る
    
    if(i <= 6){
      //var row = sheet.getRange(2,i+2,1,1).setValue("=(ArrayFormula(countif('フォーム1/2'!" + box[2*i + 1] + ':' + box[2*i + 2] + ",$B$2:$B$57)))");
      var row = sheet.getRange(2,i+2,1,1).setValue("=(ArrayFormula(countif('フォーム1/2'!" + box[2*i + 1] + ':' + box[2*i + 2] + ",$B$2:$B$"+ lastRow1+1 +")))");
      
    //}else if(i>=8 && i <= 63){
    }else if(i!=7 && i <= lastRow2/2){
      
      //var row = sheet.getRange(2,i+2,1,1).setValue("=(ArrayFormula(countif('フォーム1/2'!" + box[2*i + 1] + ':' + box[2*i + 2] + ",$B$2:$B$57)))");
      var row = sheet.getRange(2,i+2,1,1).setValue("=(ArrayFormula(countif('フォーム1/2'!" + box[2*i + 1] + ':' + box[2*i + 2] + ",$B$2:$B$"+ lastRow1+1 +")))");
      
    //}else if(64<=i && i <= 126){
    }else if(lastRow2/2<i && i <= lastRow2){
      
      //var row = sheet.getRange(2,i+2,1,1).setValue("=(ArrayFormula(countif('フォーム2/2'!" + box[2*(i-63) + 1] + ':' + box[2*(i-63) + 2] + ",$B$2:$B$57)))");
      if(lastRow2%2 == 0) var row = sheet.getRange(2,i+2,1,1).setValue("=(ArrayFormula(countif('フォーム2/2'!" + box[2*(i-lastRow2/2) + 1] + ':' + box[2*(i-lastRow2/2) + 2] + ",$B$2:$B$"+ lastRow1+1 +")))");      
      else var row = sheet.getRange(2,i+2,1,1).setValue("=(ArrayFormula(countif('フォーム2/2'!" + box[2*(i-lastRow2/2) + 2] + ':' + box[2*(i-lastRow2/2) + 3] + ",$B$2:$B$"+ lastRow1+1 +")))");      
    }
    Logger.log(box)
  }
}

//計算タブから結果タブにコピー
function myFunction2(){
  var sheet1 = ss.getSheetByName('計算');
  var sheet2 = ss.getSheetByName('結果');
  var sheet3 = ss.getSheetByName('データ入力');//
  var lastRow1 = sheet3.getRange('C69').getValue();//
  var lastRow2 = sheet3.getRange('I140').getValue();//
  
  //var data = sheet1.getRange(2,3,56,126).getValues();
  var data = sheet1.getRange(2,3,lastRow1,lastRow2).getValues();

  
  //sheet2.getRange(3,3,56,126).setValues(data);
  sheet2.getRange(3,3,lastRow1,lastRow2).setValues(data);
}

function onOpen(){
 var sheet = SpreadsheetApp.getActiveSpreadsheet();
 var entries = [{
  name: "更新",
  functionName: "myFunction2"
 }];

 sheet.addMenu("更新",entries);
};
