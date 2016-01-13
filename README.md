Query Microsoft's Exchange Web Services. Only tested on Microsoft Exchange 2010.

##Install

Install with npm:

```
npm install exchanger
```

##Module usage

###Initialize client

``` javascript
  var exchanger = require('exchanger');
  exchanger.initialize({ url: 'webmail.example.com', username: 'username', password: 'password' })
  .then(function(client) {
    console.log('Initialized!');
  });
```

###Use client

``` javascript
  exchanger.getCalendars()
  .then(function(calendars) {
    console.log(calendars);
  })
  .fail(function(error) {
    if(error.code == 401) console.log('Cant log in');
    if(error.code == 404) console.log('Cant connect to server');
    if(error.code == 'NOCLIENT') console.log('No client initialized');
  });
```

##Available methods

###exchanger.initialize(setting)

``` javascript
  exchanger.initialize({ url: 'webmail.example.com', username: 'username', password: 'password' })
  .then(function(client) {
    console.log('Initialized!');
  });
```

###exchanger.getEmails(folderName, limit)

``` javascript
  exchanger.getEmails('inbox', 50)
  .then(function(emails) {
    console.log(emails);
  });
```

###exchanger.getCalendars()

``` javascript
  exchanger.getCalendars()
  .then(function(calendars) {
    console.log(calendars);
  });
```

###exchanger.resolveNames(name)

``` javascript
  exchanger.resolveNames('username')
  .then(function(contacts) {
    console.log(contacts);
  });
```

###exchanger.getUserCalendars(folder)

``` javascript
  exchanger.getUserCalendars('calendar')
  .then(function(calendars) {
    console.log(calendars);
  });
```

###exchanger.getRootCalendar()

``` javascript
  exchanger.getRootCalendar()
  .then(function(calendar) {
    console.log(calendar);
  });
```

###exchanger.getCalendarItems()

``` javascript
  exchanger.getCalendarItems({id: 'AAAWAE1hdGhpZXUuUGVycmluQGItaS5jb20ALgAAAAAAcSKJkLtVjEmCyysMMvwm7wEASv22nSmUD0e/3dEJOBwjkwAACj31mAAA'}, '2016-01-11T00:00:00+01:00', '2016-01-17T23:59:59+01:00')
  .then(function(events) {
    console.log(events);
  });
```