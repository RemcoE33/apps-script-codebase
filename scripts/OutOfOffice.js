/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase

  1. Set you out of office inside Gmail Settings to use as template.
  2. Check GmailAPI under services inside the script editor.
  3. Run 'readCurrentOutOfOffice' and copy the HTML from the log.
  4. Change the HTML on 'responseBodyHtml' with your copied html.
  5. Now you can set 'outOfOfficeON' and 'outOfOfficeOff' on a trigger of your choose.
*/

//Global variable:
const USER = 'user@company.com'


function readCurrentOutOfOffice(){
  console.log(Gmail.Users.Settings.getVacation(USER))
}

function outOfOfficeON() {

  //Creating the dates
  const date = new Date().toISOString().split("T")[0]
  const startTime = new Date(date).getTime()
  const endTime = new Date(startTime + 172800000).getTime()

  console.log({
    date,
    startTime,
    endTime
  })

  //Build settings:
  const settings = JSON.stringify({
    restrictToContacts: false,
    enableAutoReply: true,
    responseBodyHtml: `<div dir="ltr">Hi mailer,<div><br></div><div>I&#39;m currently not in the office because i like diving with sharks. If i&#39;m not eaten alive i will get back to you first thing tomorrow!</div><div><br></div><div>Greetings,</div><div>RemcoE33</div></div>`,
    responseSubject: 'I\'m not around',
    endTime: endTime.toString(),
    startTime: startTime.toString()
  })

  const response = Gmail.Users.Settings.updateVacation(settings, USER)

  console.log({
    type: 'Turning on',
    response
  })
}

function outOfOfficeOff(){
  const settings = JSON.stringify({
    enableAutoReply: false
  })

  const response = Gmail.Users.Settings.updateVacation(settings, USER)

  console.log({
    type: 'Turning off',
    response
  })

}