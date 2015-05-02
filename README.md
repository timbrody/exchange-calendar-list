# exchange-calendar-list
Node.js module to read entries from an Exchange calendar via EWS forked from timbrody/exchange-calendar-list

This version uses node soap and xpath with xmldom to more easily navigate EWS soap messages and has been tested with Office365.

exchange-calendar-list is dependant on node-promises as the original. 

To use:
=======
var exchList = require('exchange-calendar-list')();
exchList.request({url: url, username: acc.access.username, password:acc.access.password, start: start, end : end, maximum: 50})
.then(function(items) {
  //each item contains 
  //{
  //    itemId: 'xx', changeKey: 'xx', type: 'xx', 
  //    subject: 'some subject', location: 'some place',
  //    start: 'startdate', end: 'enddate', allDay: [true or false],
  //    body: 'some extended body as text only',
  //    organizer: {name: 'organizers name', email: 'organizers email'},
  //    requiredAttendees: [{name: 'name', email: 'email@something.com'}],
  //    optionalAttendees: [{name: 'name', email: 'email@something.com'}]
  //}
}, function(err) {
  if(err) throw err;
});
