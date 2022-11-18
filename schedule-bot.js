const excelToJson = require('convert-excel-to-json');
const twilio = require('twilio');
require('dotenv').config()

const extractExcelData = excelToJson({
    sourceFile: './schedule.xlsx'
});

const findAiSchedule = (scheduleData) => {
  let ai;
  const date = Object.keys(scheduleData);
  scheduleData[date].forEach((element) => {
    if (element["A"] === "Ai") ai = element;
  });

  const arr = [];

  // Format times into an array
  for (const key in ai) {
    if (ai[key] === 'Ai') continue;
    let time = ai[key].toTimeString();
    arr.push(time);
  }

  // Loop through array of times, format
  const scheduleArr = [];
  for (let i = 0; i < arr.length-2; i++) {
    let currentTime = arr[i].slice(0, 2);
    let nextTime = arr[i+1].slice(0, 2);
    let isItAShift = arr[i+2].slice(0, 2) === "09";
    if (currentTime === '00' && nextTime === '00') {
      scheduleArr.push('OFF');
    }
    if (isItAShift) {
      if (currentTime > 12) { 
        currentTime = currentTime - 12;
        currentTime = currentTime + "PM"
      } else {
        currentTime = currentTime + "AM"
      };
      if (nextTime > 12) { 
        nextTime = nextTime - 12;
        nextTime = nextTime + "PM"
      } else {
        nextTime = nextTime + "AM"
      };

      if (currentTime[0] === "0") currentTime = currentTime.substring(1);
      if (nextTime[0] === "0") nextTime = nextTime.substring(1);
      let temp = currentTime + " - " + nextTime;
      scheduleArr.push(temp);
    }
  }

  // Link times to days of the week
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const finalSchedule = {};
  for (let x = 0; x < 7; x++) {
    finalSchedule[days[x]] = scheduleArr[x];
  }
  finalSchedule.date = date;
  return finalSchedule;
}

const scheduleObj = findAiSchedule(extractExcelData);

const createMessage = (scheduleObj) => {
  const str = `
  Hello, Ai! 🙂 This is your schedule at MUJI Yorkdale from ${scheduleObj.date}. 
  Monday: ${scheduleObj.Monday}
  Tuesday: ${scheduleObj.Tuesday}
  Wednesday: ${scheduleObj.Wednesday}
  Thursday: ${scheduleObj.Thursday}
  Friday: ${scheduleObj.Friday}
  Saturday: ${scheduleObj.Saturday}
  Sunday: ${scheduleObj.Sunday}`
  return str;
}

const messageToRelay = createMessage(scheduleObj);

const accountSid = process.env.ACCOUNT_SID;
const authToken = process.env.AUTH_TOKEN;
const client = twilio(accountSid, authToken);

client.messages 
  .create({body: messageToRelay, from: process.env.TWILIO_NUMBER, to: process.env.PARKER_PHONE_NUMBER})
  .then(message => console.log('Schedule: ', scheduleObj))
  .then(message => console.log('Message sent successfully to Parker!'))
  .catch((err) => console.log('Failure: ', err, 'The message was not send to Parker.'));        
  
client.messages 
  .create({body: messageToRelay, from: process.env.TWILIO_NUMBER, to: process.env.AI_PHONE_NUMBER})
  .then(message => console.log('Message sent successfully to Ai!'))
  .catch((err) => console.log('Failure: ', err, 'The message was not sent to Ai.')); 