// var dt = new Date().toLocaleString(undefined,{year:'numeric', month: '2-digit', day:'2-digit'}).split(' ')[0];
//     // var date  = dt.split(' ')[0];
// var startDate = dt + 'T00:00:00.0000000';
// var endDate = dt + 'T23:59:59.0000000';

// var url = '/me/calendar/calendarView?startDateTime='+ startDate + '&'+'endDateTime='+endDate;

// var url2 = '/me/calendar/calendarView?startDateTime=2017-04-21T00:00:00.0000000&endDateTime=2017-04-21T23:59:59.0000000';

// console.log(url);
// console.log(url2);

// if(url===url2){
//   console.log('yes')
// }



//expect 2017-04-21T00:00:00.0000000
// var date = new Date();



// month < 10 ? '0' + month : '' + month
// var year = date.getFullYear().toString();
// var month = ((date.getMonth() + 1) < 10 ? '0' : '') + (date.getMonth() + 1);
// var day = (date.getDate() < 10 ? '0' : '') + date.getDate();

// var startDate = year+'-'+month+'-'+day+'T'+'00:00:00.0000000';
// var endDate = year+'-'+month+'-'+day+'T'+'23:59:59.0000000';
// console.log(year, month, day);
// console.log(url);


var Moment = require('moment-timezone');
today = Moment().tz('Asian/Taipei').startOf('hour').add(24, 'hours').format('YYYY-MM-DD');
var startDate = today+'T'+'00:00:00.0000000';
var endDate = today+'T'+'23:59:59.0000000';
console.log(startDate);
console.log(typeof(startDate));

sss = Moment().startOf('hour').add(1, 'hours').format('HH:mm')
