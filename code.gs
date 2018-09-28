/*================================================================================================================*
  Classroom Attendance by Sheetgo
  ================================================================================================================
  Version:      1.0.0
  Project Page: https://github.com/Sheetgo/classroom-atenddance
  Copyright:    (c) 2018 by Sheetgo
  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  ----------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.0.0  Initial release
 *================================================================================================================*/

var ss = SpreadsheetApp.getActiveSpreadsheet();
var firstTime = PropertiesService.getUserProperties().getProperty('firstTime');

function main() {
    var classroomData = this.getDataFromClassroom();
    if (classroomData.length != 0) {
        this.clearSheets(false);
        this.insertTabValues(classroomData);
        this.clearSheets(true);
    }
}

/** Lifecycle method - When the spreadsheet is opened **/
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    if (!this.firstTime) {
        ui.createMenu('Classroom')
            .addItem('Enable classroom api', 'testClassroomApi')
            .addItem('Start', 'start')
            .addToUi();
    } else {
        ui.createMenu('Classroom')
            .addItem('Refresh', 'main')
            .addToUi();
    }
    this.enableTrigger()
}

/** Action of classroom menu - When the user clicks on start on the first time **/
function start() {
    if (!this.firstTime) {
        this.main();
        this.enableTrigger();
        PropertiesService.getUserProperties().setProperty('firstTime', true);
    }
}

/** Enable trigger daily trigger **/
function enableTrigger() {
    var trigger = ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());
    if (trigger.length > 0) {
        for (var i in trigger) {
            if (!trigger[i].getHandlerFunction() === "main") {
                ScriptApp.newTrigger('main').timeBased().everyDays(1).create();
            }
        }
    } else {
        ScriptApp.newTrigger('main').timeBased().everyDays(1).create();
    }

}

/** Checks Classroom api **/
function testClassroomApi() {
    try {
        var courses = Classroom.Courses.list()
    } catch (e) {
        if (e.message.substring(0, 38) === 'Google Classroom API has not been used') {
            var url = 'https://console.developers.google.com/apis/api/classroom.googleapis.com/overview?project=' + e.message.substring(50, 62)
            this.openUrl(url);
        }
    }
}

/** Bring organized data from Classroom **/
function getDataFromClassroom() {

    var values = [];
    var courses = this.getCourses();
    courses.map(function (course) {
        // Fills name, first name, last name and email headers
        var courseIdCollunm = ["Class"];
        var nameCollunm = ["Name"];
        var firstNameCollunm = ["First name"];
        var lastNameCollunm = ["Last name"];
        var emailCollunm = ["Email"];

        var students = this.getStudentsByCourse(course.id);
        var courseTab = [];

        // Fills name, first name, last name and email collunms
        students.map(function (student, i) {
            nameCollunm.push(student.profile.name.fullName);
            firstNameCollunm.push(student.profile.name.givenName);
            lastNameCollunm.push(student.profile.name.familyName);
            emailCollunm.push(student.profile.emailAddress);
        })

        var worksCollunm = [];
        var works = this.getCourseWorks(course.id);
        works.map(function (work) {
            var auxWorksCollunm = [];
            var auxDateWorksCollunm = [];

            // Fills all works and works delivery date headers
            auxWorksCollunm.push(work.title);
            auxDateWorksCollunm.push(work.title + " date");

            // Fills works grades by student and works delivery date
            students.map(function (student) {
                var grades = this.getGradeByWork(course.id, work.id, student.userId);
                grades.map(function (grade) {
                    if (work.dueDate) {
                        var date = work.dueDate
                        date = new Date(date.year, date.month, date.day);
                        auxDateWorksCollunm.push(date);
                    } else {
                        auxDateWorksCollunm.push(" ");
                    }
                    if (grade.draftGrade) {
                        auxWorksCollunm.push(grade.draftGrade)
                    } else {
                        auxWorksCollunm.push(" ")
                    }
                })
            })
            worksCollunm.push(auxWorksCollunm, auxDateWorksCollunm);
        })


        nameCollunm.map(function (s, i) {
            if (!courseIdCollunm.length < i) {
                courseIdCollunm.push(course.id);
            }
        })
        courseTab.push(nameCollunm);
        courseTab.push(firstNameCollunm);
        courseTab.push(lastNameCollunm);
        courseTab.push(emailCollunm);
        courseTab.push(courseIdCollunm);
        worksCollunm.map(function (work) {
            courseTab.push(work);
        })

        // Format multidimensional array to fit on sheet
        courseTab = this.transpose(courseTab);

        values.push({ tabName: course.name, tabValues: courseTab });
    })

    return values;
}

/** Get courses from Classroom **/
function getCourses() {

    try {
        var response = Classroom.Courses.list({ teacherId: 'me', courseStates: 'ACTIVE' });
        var courses = response.courses;
        var values = [];
        if (courses && courses.length > 0) {
            courses.map(function (course) {
                values.push(course)
            })
        }
        return values;
    } catch (e) {
        var ui = SpreadsheetApp.getUi();
        ui.alert("Error to get data from Classroom", JSON.stringify(e.message), ui.ButtonSet.OK);
    }
}

/** Get students by course from Classroom **/
function getStudentsByCourse(courseId) {
    var response = Classroom.Courses.Students.list(courseId);
    var students = response.students;
    var values = [];
    if (students && students.length > 0) {
        students.map(function (student) {
            values.push(student);
        })
    }
    return values;
}

/** Get grades from Classroom **/
function getGradeByWork(courseId, workId, userId) {
    var response = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, workId, { userId: userId });
    var submissions = response.studentSubmissions;
    var values = [];
    if (submissions && submissions.length > 0) {
        submissions.map(function (submission) {
            values.push(submission);
        })
    }
    return values;
}

/** Get course works from Classroom **/
function getCourseWorks(courseId) {
    var response = Classroom.Courses.CourseWork.list(courseId);
    var works = response.courseWork;
    var values = [];
    if (works && works.length > 0) {
        works.map(function (work) {
            values.push(work);
        })
    }
    return values;
}

/** Clear spreadsheet **/
function clearSheets(temp) {
    if (temp) {
        var sheet = ss.getSheetByName('temp')
        ss.deleteSheet(sheet);
    } else {
        ss.insertSheet('temp')
        this.ss.getSheets().map(function (sheet) {
            var name = sheet.getSheetName();
            if (name != 'temp')
                ss.deleteSheet(sheet);
        })
    }
}

/** Insert tabs data on spreadsheet **/
function insertTabValues(tabs) {
    tabs.map(function (t) {
        var tab = ss.insertSheet(t.tabName);
        tab.getRange(1, 1, t.tabValues.length, t.tabValues[0].length).setValues(t.tabValues);
    })
}


/** Transpose data from multidimentsional array **/
function transpose(a) {
    var w = a.length || 0;
    var h = a[0] instanceof Array ? a[0].length : 0;

    if (h === 0 || w === 0) { return []; }
    var i, j, t = [];

    for (i = 0; i < h; i++) {
        t[i] = [];
        for (j = 0; j < w; j++) {
            t[i][j] = a[j][i];
        }
    }
    return t;
}

/** Method to open external link on new tab **/
function openUrl(url) {
    var ui = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutput('<html><script>'
        + 'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
        + 'var a = document.createElement("a"); a.href="' + url + '"; a.target="_blank";'
        + 'if(document.createEvent){'
        + '  var event=document.createEvent("MouseEvents");'
        + '  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'
        + '  event.initEvent("click",true,true); a.dispatchEvent(event);'
        + '}else{ a.click() }'
        + 'close();'
        + '</script>'
        // Offer URL as clickable link in case above code fails.
        + '<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="' + url + '" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
        + '<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
        + '</html>')
        .setWidth(90).setHeight(1);
    ui.showModalDialog(html, "Opening ...");
}

