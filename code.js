/**
 * source references
 * index to col name : https://gist.github.com/rheajt/d48a54be3aad01a1931a1a433dc99e5c
 * lunar to sol calendar : https://gist.github.com/twkang/6c2979c9d0bc7431792e29181f0febff
 */

// following two import lines should be removed in .gs file @Google Sheet > Extensions > Apps Script
import 'google-apps-script';
import 'keys.js';
/**
 * following static variable assigned @/keys.js
 * KASI_KEY_ENCODED
 * KASI_KEY_DECODED
 * CALENDAR_ID
 */

function update_anniversary() {
  var eventCal = CalendarApp.getCalendarById(CALENDAR_ID);

  var spreadsheet = SpreadsheetApp.getActive().getSheetByName("Birthday");
  do_update(spreadsheet, eventCal, ["🎂", "'s", "Birthday"]);
  console.log("B-day done");

  var spreadsheet = SpreadsheetApp.getActive().getSheetByName("Welcome day");
  do_update(spreadsheet, eventCal, ["💐", "'s", "Welcome day"]);
  console.log("W-day done");
}

function do_update(spreadsheet, eventCal, title_ingredients) {
  
  var row_max = spreadsheet.getMaxRows();
  var col_max = spreadsheet.getMaxColumns();
  var col_max_letter = convert_index_to_letter(col_max);
  var signups = spreadsheet.getRange(`A2:${col_max_letter}${row_max}`).getValues();
  
  var genesis = new Date(`${start_yyyy}-01-01 09:00:00`);
  var today = new Date()
  var limit_yyyy = today.getFullYear() + 1;
  var apocalypse = new Date(`${limit_yyyy}-12-31 23:59:59`);

  var event_ids = [];

  for (x=0; x<signups.length; x++) {
    var shift = signups[x];
    
    var name = shift[0];
    if ( name )
    {
      var title = [title_ingredients[0], name, title_ingredients[1], title_ingredients[2]].join(" ");

      var startDate = shift[1];
      if (startDate)
      {
        var is_lunar = shift[2];

        startDate = new Date(startDate);
        startDate = new Date(startDate.getTime() + 24 * 60 * 60 * 1000);
        var origin_yyyy = startDate.getFullYear();
        var origin_mm = startDate.getMonth() + 1;
        var origin_dd = startDate.getDate();

        var start_yyyy = 2020;
        if (title_ingredients[2] === "Welcome day")
        {
          start_yyyy = origin_yyyy;
        }

        var events = eventCal.getEvents(genesis, apocalypse, {search: name});
        events = events.filter( event => event.getSummary().search(title_ingredients[2]) !== -1 );
        if (events.length !== 0) // target member's event exist -----------------------------------------------------------------------------------------------------------
        {
          console.log( `${x} : ${title} (${origin_yyyy}-${origin_mm}-${origin_dd}) : event already exists`);
          var lastest_yyyy;
          events.forEach( function (event) {
            event_ids.push( event.getId() );

            var old_birthday = event.getAllDayStartDate();
            var old_yyyy = lastest_yyyy = old_birthday.getFullYear();
            var old_mm = old_birthday.getMonth() + 1;
            var old_dd = old_birthday.getDate();
            
            if ( is_lunar === 1)
            {
              console.log( `apply lunar to sol calendar`);
              converted = lun2sol(old_yyyy, origin_mm, origin_dd, 1);
              target_mm = parseInt(converted.split("-")[1]);
              target_dd = parseInt(converted.split("-")[2]);
            }
            else
            {
              target_mm = origin_mm;
              target_dd = origin_dd;
            }
            var new_birthday = new Date(`${old_yyyy}-${target_mm}-${target_dd} 09:00:00`);

            if ( `${old_mm}-${old_dd}` !== `${target_mm}-${target_dd}` )
            {
              console.log("update birthday");
              console.log("old yyyy-mm-dd : " + `${old_yyyy}-${old_mm}-${old_dd}`);
              console.log("new yyyy-mm-dd : " + `${old_yyyy}-${target_mm}-${target_dd}`);
              event.setAllDayDate(new_birthday);
              console.log(`update event : ${old_yyyy}-${target_mm}-${target_dd}`);
            }
          });

          console.log("lastest_yyyy : " + lastest_yyyy );
          if ( lastest_yyyy < limit_yyyy ) for ( var target_yyyy = lastest_yyyy + 1; target_yyyy < limit_yyyy + 1; target_yyyy++)
          {
            if ( is_lunar === 1)
            {
              console.log( `apply lunar to sol calendar`);
              converted = lun2sol(target_yyyy, origin_mm, origin_dd, 1);
              target_mm = parseInt(converted.split("-")[1]);
              target_dd = parseInt(converted.split("-")[2]);
            }
            else
            {
              target_mm = origin_mm;
              target_dd = origin_dd;
            }

            if (title_ingredients[2] === "Birthday")
            {
              var desp = `${title_ingredients[0]} Happy Birthday ${name}, Congrats! `;
            }
            else if (title_ingredients[2] === "Welcome day")
            {
              var desp = `${title_ingredients[0]} Thank you ${name}! It's ${convert_int_to_ordinal(target_yyyy - origin_yyyy)} Anniversary!`;
            }
            else
            {
              var desp = "";
            }
            var event = eventCal.createAllDayEvent(title, new Date(`${target_yyyy}-${target_mm}-${target_dd} 09:00:00`), {description: desp} );
            event_ids.push( event.getId() );
            console.log(`create event : ${target_yyyy}-${target_mm}-${target_dd}`);
          }
        }
        else // target member's event NOT exist -----------------------------------------------------------------------------------------------------------
        {
          console.log( `${x} : ${title} (${origin_yyyy}-${origin_mm}-${origin_dd}) : no event exist. create a new one from ${start_yyyy} to ${limit_yyyy}`);
          for ( var target_yyyy = start_yyyy; target_yyyy < limit_yyyy + 1; target_yyyy++)
          {
            if ( is_lunar === 1)
            {
              console.log( `apply lunar to sol calendar`);
              converted = lun2sol(target_yyyy, origin_mm, origin_dd, 1);
              target_mm = parseInt(converted.split("-")[1]);
              target_dd = parseInt(converted.split("-")[2]);
            }
            else
            {
              target_mm = origin_mm;
              target_dd = origin_dd;
            }

            if (title_ingredients[2] === "Birthday")
            {
              var desp = `${title_ingredients[0]} Happy Birthday ${name}, Congrats! `;
            }
            else if (title_ingredients[2] === "Welcome day")
            {
              var work_years = target_yyyy - origin_yyyy;
              if (work_years === 0)
              {
                var desp = `${title_ingredients[0]} Welcome ${name}! It's ${convert_int_to_ordinal(work_years)} Anniversary!`;
              }
              else
              {
                var desp = `${title_ingredients[0]} Thank you ${name}! It's ${convert_int_to_ordinal(work_years)} Anniversary!`;
              }
            }
            else
            {
              var desp = "";
            }
            var event = eventCal.createAllDayEvent(title, new Date(`${target_yyyy}-${target_mm}-${target_dd} 09:00:00`), {description: desp} );
            event_ids.push( event.getId() );
            console.log(`create event : ${target_yyyy}-${target_mm}-${target_dd}`);
          }
        }

        // check duplicated event -----------------------------------------------------------------------------------------------------------
        console.log("check duplicated events");
        var all_events = eventCal.getEvents(genesis, apocalypse, {search: title_ingredients[2]});
        for ( var target_yyyy = start_yyyy; target_yyyy < limit_yyyy + 1; target_yyyy++)
        {
          var target_events = all_events.filter( event => event.getAllDayStartDate().getFullYear() === target_yyyy && event.getSummary().search(name) !== -1 );
          if ( target_events.length > 1 )
          {
            for ( var i = 0; i < target_events.length - 1; i++ )
            {
              var target_event = target_events[i];
              target_date = target_event.getAllDayStartDate();
              console.log(`delete duplicated event : ${target_event.getSummary()} (${target_date.getFullYear()}-${target_date.getMonth()+1}-${target_date.getDate()})`);
              target_event.deleteEvent();
            }
          }
        }
      }
      else
      {
        console.log( x + " : " + title + " : no startDate");
      }
    }
  }

  // check deleted member event -----------------------------------------------------------------------------------------------------------
  console.log("check deleted memeber events");
  all_events = eventCal.getEvents(genesis, apocalypse, {search: title_ingredients[2]});
  all_events.forEach(event => {
    if ( !event_ids.includes( event.getId() ) )
    {
      var target_date = event.getAllDayStartDate();
      console.log(`delete ${event.getSummary()} (${target_date.getFullYear()}-${target_date.getMonth()+1}-${target_date.getDate()})`);
      event.deleteEvent();
    }
  });

}

function convert_index_to_letter(col_index) {
   var temp, letter = '';
   while (col_index > 0) {
      temp = (col_index - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      col_index = (col_index - temp - 1) / 26;
   }
   return letter;
}

function convert_int_to_ordinal(n) {
  if ( n >= 11 && n <=13 )
  {
    var suffix = "th";
  }
  else
  {
    switch (n % 10) {
      case 1:
        var suffix = "st";
        break;
      case 2:
        var suffix = "nd";
        break;
      case 3:
        var suffix = "rd";
        break;
      default:
        var suffix = "th";
    }
  }

  return `${n}${suffix}`;
}

/**
 * 음양력변환 함수 (한국천문연구원 Open API 이용)
 * 키발급: https://www.data.go.kr/dataset/15012679/openapi.do
 * 발급받은 키로 KASI_KEY_ENCODED 부분 교체
 * source reference : https://gist.github.com/twkang/6c2979c9d0bc7431792e29181f0febff
 */

/**
 * 음력 10월 3일의 2013년 양력일자를 알고 싶은 경우 lun2sol("2013", "10", "03", 1) 으로 호출
 * 응답으로는 '2013-11-05' 와 같이 스트링을 리턴.
 * yoon (윤달구분) : 1 - 평달, 2 - 윤달
 */
function lun2sol(yyyy, mm, dd, yoon) {

  if (typeof yyyy !== "string") yyyy = yyyy.toString();
  if (typeof mm !== "string") mm = mm.toString();
  if (typeof dd !== "string") dd = dd.toString();
  mm = mm.padStart(2, "0");
  dd = dd.padStart(2, "0");

  /* 이미 조회했던 일자이면 ScriptProperties 에서 읽어 온다. -- 네트웍을 통한 조회는 한번만 하도록 */  
  var prop = ScriptProperties.getProperty("lun2sol/" + yyyy + "/" + mm + "/" + dd + "/" + yoon);
  if (prop != null) {
    Logger.log("Read from db: " + prop);
    return prop;
  }
  
  /* Calendar API - convert lunar to solar calendar */
  var param = `lunYear=${yyyy}` + `&lunMonth=${mm}` + `&lunDay=${dd}` + 
    `&ServiceKey=${KASI_KEY_ENCODED}` +
    "&_type=json";
  var url = "http://apis.data.go.kr/B090041/openapi/service/LrsrCldInfoService/getSolCalInfo?" + param;

  var result = UrlFetchApp.fetch(url);
  var data = JSON.parse(result.getContentText());
  
  var itemCount = data.response.body.totalCount;

  if (itemCount == 1) {
    var item = data.response.body.items.item;

    var new_prop = item.solYear + "-" + item.solMonth + "-" + item.solDay
    ScriptProperties.setProperty("lun2sol/" + yyyy + "/" + mm + "/" + dd + "/" + yoon, new_prop);
    
    return new_prop;
  }
  else if (itemCount > 1) {
    data.response.body.items.item.forEach( item => {
      if ((item.lunLeapmonth == "평") && (yoon == 1)) {
        var new_prop = item.solYear + "-" + item.solMonth + "-" + item.solDay
        ScriptProperties.setProperty("lun2sol/" + yyyy + "/" + mm + "/" + dd + "/" + yoon, new_prop); 
        return new_prop;
      }
      else if ((item.lunLeapmonth == "윤") && (yoon == 2)) {
        var new_prop = item.solYear + "-" + item.solMonth + "-" + item.solDay
        ScriptProperties.setProperty("lun2sol/" + yyyy + "/" + mm + "/" + dd + "/" + yoon, new_prop); 
        return new_prop;
      }
    });
  }
  
  return "ERROR";
}

/**
 * 양력 2013년 1월 26일의 음력일자를 알고 싶은 경우 sol2lun("2013", "01", "26") 으로 호출
 * 응답으로는 '2012-12-15 평' 과 같이 스트링을 리턴.
 */
function sol2lun(yyyy, mm, dd) {
  if (typeof yyyy !== "string") yyyy = yyyy.toString();
  if (typeof mm !== "string") mm = mm.toString();
  if (typeof dd !== "string") dd = dd.toString();
  mm = mm.padStart(2, "0");
  dd = dd.padStart(2, "0");

  /* 이미 조회했던 일자이면 ScriptProperties 에서 읽어 온다. -- 네트웍을 통한 조회는 한번만 하도록 */  
  var prop = ScriptProperties.getProperty("sol2lun/" + yyyy + "/" + mm + "/" + dd);
  if (prop != null) {
    Logger.log("Read from db: " + prop);
    return prop;
  }

  /* Calendar API - convert solar to lunar calendar */
  var param = `solYear=${yyyy}` + `&solMonth=${mm}` + `&solDay=${dd}` + 
    `&ServiceKey=${KASI_KEY_ENCODED}` +
    "&_type=json";
  var url = "http://apis.data.go.kr/B090041/openapi/service/LrsrCldInfoService/getLunCalInfo?" + param;

  var result = UrlFetchApp.fetch(url);
  var data = JSON.parse(result.getContentText());
  
  var itemCount = data.response.body.totalCount;
  
  if (itemCount == 1) {
    var item = data.response.body.items.item;

    yoonstr = item.lunLeapmonth;
    var new_prop = item.lunYear + "-" + item.lunMonth + "-" + item.lunDay + " " + yoonstr;
    ScriptProperties.setProperty("sol2lun/" + yyyy + "/" + mm + "/" + dd, new_prop);

    return new_prop;
  }
  return "ERROR";
}