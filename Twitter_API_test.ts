import RPA from 'ts-rpa';
const fs = require('fs');

// 読み込みする スプレッドシートID と シート名 の記載
const SSID = process.env.Twitter_SheetID_test;
const SSName1 = process.env.Twitter_SheetName_test_1;
const SSName2 = process.env.Twitter_SheetName_test_2;

var Twitter = require('twitter');
var client = new Twitter({
  consumer_key: process.env.con_key,
  consumer_secret: process.env.con_secret,
  access_token_key: process.env.token_key,
  access_token_secret: process.env.token_secret
});

// Twitter で検索するキーワード / 検索数
const KeyWord = [''];
const Count = [10];

async function Start() {
  // 実行前にダウンロードフォルダを全て削除する
  await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
  // デバッグログを最小限にする
  RPA.Logger.level = 'INFO';
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  const TextData = [];
  const Max_id = [''];
  // スプレッドシートから 検索キーワード・過去ID を取得。作業ステータスを変更する
  await GetTweetID(KeyWord, Max_id);
  // Twitter で検索して、データを取得 / TextData に格納する
  await GetTwitterData(TextData, KeyWord, Max_id);
  await RPA.sleep(3000);
  // Max_id を取得する
  await RPA.Logger.info(TextData.slice(-1)[0][6]);
  const LastMax_id = await TextData.slice(-1)[0][6];
  // スプレッドシートにデータを貼り付ける
  await DataPaste(TextData, LastMax_id);
}

Start();

// スプレッドシートから tweetID を取得する
async function GetTweetID(KeyWord, Max_id) {
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!D3:H4`
  });
  await RPA.Logger.info(FirstData);
  const TextValues = [];
  const remainingAPI = Number(FirstData[0][1]) - 1;
  TextValues.push('作業中...');
  TextValues.push(String(remainingAPI));
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!D3:E4`,
    values: [TextValues]
  });
  // 取得件数
  if (FirstData[0][2] != undefined) {
    Count[0] = Number(FirstData[0][2]);
  }
  // 検索キーワード
  if (FirstData[0][3] != undefined) {
    KeyWord[0] = String(FirstData[0][3]);
  }
  // Max_id
  if (FirstData[0][4] != undefined) {
    // Max_id より 1低い数値を取得する
    const Max_id_low1 = Number(FirstData[0][4]) - 10;
    Max_id[0] = String(Max_id_low1);
  }
  await RPA.Logger.info(`検索キーワード → ${KeyWord}`);
  await RPA.Logger.info(`検索する Max_id → ${Max_id}`);
}

// Twitterから textやIDを取得する
async function GetTwitterData(TextData, KeyWord, Max_id) {
  await client.get(
    'search/tweets',
    { q: `${KeyWord}`, lang: 'ja', count: Count[0], max_id: Max_id[0] },
    function(error, tweets, response) {
      tweets.statuses.forEach(function(tweet) {
        //console.log(tweet);
        //console.log(tweet.id_str);
        const firstlist = [];
        firstlist.push(
          tweet.created_at,
          tweet.text,
          tweet.user.screen_name,
          tweet.user.name,
          String(tweet.user.followers_count),
          tweet.user.created_at,
          tweet.id_str
        );
        TextData.push(firstlist);
      });
    }
  );
}

// スプレッドシートにデータを貼り付ける
async function DataPaste(TextData, LastMax_id) {
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!B1:B9000`
  });
  let Row = '';
  const RowNumber = Number(FirstData.length + 1);
  Row = String(RowNumber);
  await RPA.Logger.info(`この行から貼り付けします → ${Row}`);
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!A${Row}:G9000`,
    values: TextData
  });
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!D3:D3`,
    values: [['完了']]
  });
  // H3 に MaxID を記載する
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName2}!H3:H3`,
    values: [[LastMax_id]]
  });
}
