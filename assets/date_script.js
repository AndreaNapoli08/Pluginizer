// licenza d'uso riservata ad Andrea Napoli e all'università si Bologna

function CalendarControl() {
    const calendar = new Date();

    // Ottieni i parametri dalla query string dell'URL
    const queryParams = new URLSearchParams(window.location.search);

    // Leggi il valore di 'selectedText' dalla query string
    const information = queryParams.get('information');
    const parsedInformation = JSON.parse(information);
    if(parsedInformation != null){
        document.getElementById("timePicker").value = parsedInformation.time;
    }
    

    const calendarControl = {
        localDate: new Date(),
        prevMonthLastDate: null,
        calWeekDays: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
        calMonthName: [
            "Jan",
            "Feb",
            "Mar",
            "Apr",
            "May",
            "Jun",
            "Jul",
            "Aug",
            "Sep",
            "Oct",
            "Nov",
            "Dec"
        ],

        saveData: function () {
            const selectedDay = calendarControl.selectedDate.getDate();
            const selectedMonth = calendarControl.calMonthName[calendarControl.selectedDate.getMonth()];
            const selectedYear = calendarControl.selectedDate.getFullYear();

            // Recupera l'ora selezionata
            const selectedTime = document.getElementById("timePicker").value;

            // Crea un oggetto contenente i dati da inviare all'add-in
            const dataToSend = {
                entity: "date",
                day: selectedDay,
                month: selectedMonth,
                year: selectedYear,
                time: selectedTime,
            };
            Office.onReady(function (info) {
                if (info.host === Office.HostType.Word || info.host === Office.HostType.Excel || info.host === Office.HostType.PowerPoint) {
                    // Invia i dati all'add-in utilizzando Office.context.ui.messageParent
                    Office.context.ui.messageParent(JSON.stringify(dataToSend));
                } else {
                    // Se l'add-in non è eseguito in un ambiente Office corretto, gestisci il caso di errore
                    console.log("Errore: ambiente Office non riconosciuto");
                }
            });
        },

        daysInMonth: function (month, year) {
            return new Date(year, month, 0).getDate();
        },
        firstDay: function () {
            return new Date(calendar.getFullYear(), calendar.getMonth(), 1);
        },
        lastDay: function () {
            return new Date(calendar.getFullYear(), calendar.getMonth() + 1, 0);
        },
        firstDayNumber: function () {
            return calendarControl.firstDay().getDay() + 1;
        },
        lastDayNumber: function () {
            return calendarControl.lastDay().getDay() + 1;
        },
        getPreviousMonthLastDate: function () {
            let lastDate = new Date(
                calendar.getFullYear(),
                calendar.getMonth(),
                0
            ).getDate();
            return lastDate;
        },
        navigateToPreviousMonth: function () {
            calendar.setMonth(calendar.getMonth() - 1);
            calendarControl.attachEventsOnNextPrev();
        },
        navigateToNextMonth: function () {
            calendar.setMonth(calendar.getMonth() + 1);
            calendarControl.attachEventsOnNextPrev();
        },
        navigateToCurrentMonth: function () {
            let currentMonth = calendarControl.localDate.getMonth();
            let currentYear = calendarControl.localDate.getFullYear();
            calendar.setMonth(currentMonth);
            calendar.setYear(currentYear);
            calendarControl.attachEventsOnNextPrev();
        },
        displayYear: function () {
            let yearLabel = document.querySelector(".calendar .calendar-year-label");
            yearLabel.innerHTML = calendar.getFullYear();
        },
        displayMonth: function () {
            let monthLabel = document.querySelector(
                ".calendar .calendar-month-label"
            );
            monthLabel.innerHTML = calendarControl.calMonthName[calendar.getMonth()];
        },
        selectDate: function (e) {
            console.log(
                `${e.target.textContent} ${calendarControl.calMonthName[calendar.getMonth()]
                } ${calendar.getFullYear()}`
            );
        },

        selectDate: function (e) {
            const selectedDay = e.target.textContent;
            const selectedMonth = calendarControl.calMonthName[calendar.getMonth()];
            const selectedYear = calendar.getFullYear();

            // Aggiorna la data selezionata
            calendarControl.selectedDate = new Date(selectedYear, calendar.getMonth(), selectedDay);

            // Rimuovi la classe "calendar-selected" da tutti gli elementi ".number-item"
            const dateItems = document.querySelectorAll(".calendar .number-item");
            dateItems.forEach((item) => item.classList.remove("calendar-selected"));

            // Aggiungi la classe "calendar-selected" solo all'elemento corrispondente al giorno selezionato
            e.target.parentElement.classList.add("calendar-selected");

            // Visualizza la data selezionata nel div ".calendar-today-date"
            document.querySelector(".calendar .calendar-today-date").textContent =
                "Selected: " +
                calendarControl.calWeekDays[calendarControl.selectedDate.getDay()] +
                ", " +
                calendarControl.selectedDate.getDate() +
                " " +
                calendarControl.calMonthName[calendarControl.selectedDate.getMonth()] +
                " " +
                calendarControl.selectedDate.getFullYear();
        },

        plotSelectors: function () {
            document.querySelector(
                ".calendar"
            ).innerHTML += `<div class="calendar-inner"><div class="calendar-controls">
          <div class="calendar-prev"><a href="#"><svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128"><path fill="#666" d="M88.2 3.8L35.8 56.23 28 64l7.8 7.78 52.4 52.4 9.78-7.76L45.58 64l52.4-52.4z"/></svg></a></div>
          <div class="calendar-year-month">
          <div class="calendar-month-label"></div>
          <div>-</div>
          <div class="calendar-year-label"></div>
          </div>
          <div class="calendar-next"><a href="#"><svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128"><path fill="#666" d="M38.8 124.2l52.4-52.42L99 64l-7.77-7.78-52.4-52.4-9.8 7.77L81.44 64 29 116.42z"/></svg></a></div>
          </div>
          <div class="calendar-today-date"></div>
          <div class="calendar-body"></div>
          <div class="bottom-right-button"> <!-- Move the button element here -->
              <button>Save</button>
          </div>
          </div>`;
            const saveButton = document.querySelector(".calendar .bottom-right-button button");
            saveButton.addEventListener("click", calendarControl.saveData); // Attach the event listener after creating the button
        },
        plotDayNames: function () {
            for (let i = 0; i < calendarControl.calWeekDays.length; i++) {
                document.querySelector(
                    ".calendar .calendar-body"
                ).innerHTML += `<div>${calendarControl.calWeekDays[i]}</div>`;
            }
        },
        plotDates: function () {
            document.querySelector(".calendar .calendar-body").innerHTML = "";
            calendarControl.plotDayNames();
            calendarControl.displayMonth();
            calendarControl.displayYear();
            let count = 1;
            let prevDateCount = 0;

            calendarControl.prevMonthLastDate = calendarControl.getPreviousMonthLastDate();
            let prevMonthDatesArray = [];
            let calendarDays = calendarControl.daysInMonth(
                calendar.getMonth() + 1,
                calendar.getFullYear()
            );
            // dates of current month
            for (let i = 1; i < calendarDays; i++) {
                if (i < calendarControl.firstDayNumber()) {
                    prevDateCount += 1;
                    document.querySelector(
                        ".calendar .calendar-body"
                    ).innerHTML += `<div class="prev-dates"></div>`;
                    prevMonthDatesArray.push(calendarControl.prevMonthLastDate--);
                } else {
                    document.querySelector(
                        ".calendar .calendar-body"
                    ).innerHTML += `<div class="number-item" data-num=${count}><a class="dateNumber" href="#">${count++}</a></div>`;
                }
            }
            //remaining dates after month dates
            for (let j = 0; j < prevDateCount + 1; j++) {
                document.querySelector(
                    ".calendar .calendar-body"
                ).innerHTML += `<div class="number-item" data-num=${count}><a class="dateNumber" href="#">${count++}</a></div>`;
            }
            calendarControl.plotPrevMonthDates(prevMonthDatesArray);
            calendarControl.plotNextMonthDates();
        },
        attachEvents: function () {
            let prevBtn = document.querySelector(".calendar .calendar-prev a");
            let nextBtn = document.querySelector(".calendar .calendar-next a");
            let todayDate = document.querySelector(".calendar .calendar-today-date");
            let dateNumber = document.querySelectorAll(".calendar .dateNumber");
            prevBtn.addEventListener(
                "click",
                calendarControl.navigateToPreviousMonth
            );
            nextBtn.addEventListener("click", calendarControl.navigateToNextMonth);
            todayDate.addEventListener(
                "click",
                calendarControl.navigateToCurrentMonth
            );
            for (var i = 0; i < dateNumber.length; i++) {
                dateNumber[i].addEventListener(
                    "click",
                    calendarControl.selectDate,
                    false
                );
            }
        },
        plotPrevMonthDates: function (dates) {
            dates.reverse();
            for (let i = 0; i < dates.length; i++) {
                if (document.querySelectorAll(".prev-dates")) {
                    document.querySelectorAll(".prev-dates")[i].textContent = dates[i];
                }
            }
        },
        plotNextMonthDates: function () {
            let childElemCount = document.querySelector('.calendar-body').childElementCount;
            //7 lines
            if (childElemCount > 42) {
                let diff = 49 - childElemCount;
                calendarControl.loopThroughNextDays(diff);
            }

            //6 lines
            if (childElemCount > 35 && childElemCount <= 42) {
                let diff = 42 - childElemCount;
                calendarControl.loopThroughNextDays(42 - childElemCount);
            }

        },
        loopThroughNextDays: function (count) {
            if (count > 0) {
                for (let i = 1; i <= count; i++) {
                    document.querySelector('.calendar-body').innerHTML += `<div class="next-dates">${i}</div>`;
                }
            }
        },
        attachEventsOnNextPrev: function () {
            calendarControl.plotDates();
            calendarControl.attachEvents();
        },
        init: function () {
            let month;
            console.log(parsedInformation)
            if (parsedInformation != null) {
                switch (parsedInformation.month) {
                    case "Jan":
                        month = 0;
                        break;
                    case "Feb":
                        month = 1;
                        break;
                    case "Mar":
                        month = 2;
                        break;
                    case "Apr":
                        month = 3;
                        break;
                    case "May":
                        month = 4;
                        break;
                    case "Jun":
                        month = 5;
                        break;
                    case "Jul":
                        month = 6;
                        break;
                    case "Aug":
                        month = 7;
                        break;
                    case "Sep":
                        month = 8;
                        break;
                    case "Oct":
                        month = 9;
                        break;
                    case "Nov":
                        month = 10;
                        break;
                    case "Dec":
                        month = 11;
                        break;
                }
                calendar.setFullYear(parsedInformation.year, month, parsedInformation.day);
            }
            calendarControl.plotSelectors();
            calendarControl.plotDates();
            calendarControl.attachEvents();
            if(parsedInformation != null){
                const dateNumberItems = document.querySelectorAll(".calendar .dateNumber");
                dateNumberItems.forEach((item) => {
                    const day = parseInt(item.textContent);
                    if (
                        day === parsedInformation.day &&
                        calendar.getMonth() === month &&
                        calendar.getFullYear() === parsedInformation.year
                    ) {
                        item.parentElement.classList.add("calendar-selected");
                    }
                });
            }
        }
    };
    calendarControl.init();
}

const calendarControl = new CalendarControl();
