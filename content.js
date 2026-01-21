(function () {
  "use strict";

  const EXTENSION_PREFIX = "mq-export";

  let popoverOpen = false;
  let exportButton = null;
  let popover = null;
  let overlay = null;

  function formatDateForFilename(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}${month}${day}`;
  }

  function formatDateTimeForICS(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    const seconds = String(date.getSeconds()).padStart(2, "0");
    return `${year}${month}${day}T${hours}${minutes}${seconds}`;
  }

  /**
   * Generate a unique ID for calendar events
   */
  function generateUID() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}@mq-timetable-exporter`;
  }

  /**
   * Escape special characters for iCalendar format
   */
  function escapeICSText(text) {
    if (!text) return "";
    return text
      .replace(/\\/g, "\\\\")
      .replace(/;/g, "\\;")
      .replace(/,/g, "\\,")
      .replace(/\n/g, "\\n");
  }

  function parseDateTime(dateStr, timeStr) {
    let date;

    if (dateStr.includes("/")) {
      const [day, month, year] = dateStr.split("/");
      date = new Date(year, month - 1, day);
    } else if (dateStr.includes("-")) {
      date = new Date(dateStr);
    } else {
      date = new Date(dateStr);
    }

    if (timeStr) {
      const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
      if (timeMatch) {
        let hours = parseInt(timeMatch[1], 10);
        const minutes = parseInt(timeMatch[2], 10);
        const isPM = timeMatch[3] && timeMatch[3].toUpperCase() === "PM";

        if (isPM && hours < 12) hours += 12;
        if (!isPM && hours === 12) hours = 0;

        date.setHours(hours, minutes, 0, 0);
      }
    }

    return date;
  }

  function extractTimetableData(baseDate) {
    const events = [];

    // Build a map of day number to actual calendar date
    const dayNumberToDate = {};
    const dayHeaders = document.querySelectorAll(
      ".t1cal-day-header[data-t1-day-number]",
    );

    // Parse the calendar's month/year from the header
    const monthAndDateSpan = document.querySelector(".monthAndDate");
    const yearSpan = monthAndDateSpan
      ? monthAndDateSpan.nextElementSibling
      : null;

    if (monthAndDateSpan && yearSpan) {
      const dateText = monthAndDateSpan.textContent.trim();
      const yearText = yearSpan.textContent.trim();
      const monthMatch = dateText.match(/(\w+)/);

      if (monthMatch) {
        const monthName = monthMatch[1];
        const months = {
          Jan: 0,
          Feb: 1,
          Mar: 2,
          Apr: 3,
          May: 4,
          Jun: 5,
          Jul: 6,
          Aug: 7,
          Sep: 8,
          Oct: 9,
          Nov: 10,
          Dec: 11,
        };
        const monthNum = months[monthName];
        const year = parseInt(yearText, 10);

        if (monthNum !== undefined && !isNaN(year)) {
          // Build map from each day header
          dayHeaders.forEach((header) => {
            const dayNum = parseInt(
              header.getAttribute("data-t1-day-number"),
              10,
            );
            const dateElem = header.querySelector(".t1cal-day-header-date");
            if (dateElem) {
              const dayOfMonth = parseInt(dateElem.textContent.trim(), 10);
              dayNumberToDate[dayNum] = new Date(year, monthNum, dayOfMonth);
            }
          });
        }
      }
    }

    // Find all calendar days and extract events from each
    const calendarDays = document.querySelectorAll(
      ".t1cal-day[data-t1-day-number]",
    );

    calendarDays.forEach((dayDiv) => {
      const dayNum = parseInt(dayDiv.getAttribute("data-t1-day-number"), 10);
      let eventDate = dayNumberToDate[dayNum];

      if (!eventDate) {
        return;
      }

      // For events before the start date, check if we need to shift them to a future week
      if (eventDate < baseDate) {
        // This is a recurring event (or might be)
        // Calculate the next occurrence of this day >= baseDate
        const daysUntilThisDay = (dayNum - baseDate.getDay() + 7) % 7;
        const nextOccurrence = new Date(baseDate);
        nextOccurrence.setDate(
          nextOccurrence.getDate() +
            (daysUntilThisDay === 0 ? 0 : daysUntilThisDay),
        );
        nextOccurrence.setHours(0, 0, 0, 0);
        eventDate = nextOccurrence;
      }

      const thumbnailItems = dayDiv.querySelectorAll(".thumbnailItem");

      thumbnailItems.forEach((item) => {
        try {
          const event = extractEventFromThumbnail(item, eventDate, baseDate);
          if (event) {
            events.push(event);
          }
        } catch (e) {
          // Silently skip failed extractions
        }
      });
    });

    return events;
  }

  function extractEventFromThumbnail(thumbnailItem, eventDate, baseDate) {
    const courseCodeField = thumbnailItem.querySelector(
      ".thbFld_spkStudyPackageCode .editorField",
    );
    const activityNameField = thumbnailItem.querySelector(
      ".thbFld_Description1 .editorField",
    );
    const eventTimeDiv = thumbnailItem.querySelector(".eventTime");
    const locationField = thumbnailItem.querySelector(
      ".thbFld_SubHeading2 .editorField",
    );
    const recurrenceField = thumbnailItem.querySelector(
      ".thbFld_SubHeading1 .editorField",
    );

    if (!eventTimeDiv) {
      return null;
    }

    const courseCode = courseCodeField
      ? courseCodeField.textContent.trim()
      : "";
    const activityName = activityNameField
      ? activityNameField.textContent.trim()
      : "";
    const location = locationField ? locationField.textContent.trim() : "";
    const recurrence = recurrenceField
      ? recurrenceField.textContent.trim()
      : "";

    const timeText = eventTimeDiv.textContent.trim();

    const timeMatch = timeText.match(
      /(\d{1,2}):(\d{2})([ap]m)?\s*-\s*(\d{1,2}):(\d{2})([ap]m)?/i,
    );

    if (!timeMatch) {
      return null;
    }

    let startHour = parseInt(timeMatch[1], 10);
    const startMin = parseInt(timeMatch[2], 10);
    let startPeriod = (timeMatch[3] || "").toLowerCase();

    let endHour = parseInt(timeMatch[4], 10);
    const endMin = parseInt(timeMatch[5], 10);
    let endPeriod = (timeMatch[6] || "").toLowerCase();

    if (startPeriod === "pm" && startHour < 12) startHour += 12;
    if (startPeriod === "am" && startHour === 12) startHour = 0;

    if (endPeriod === "pm" && endHour < 12) endHour += 12;
    if (endPeriod === "am" && endHour === 12) endHour = 0;

    if (!endPeriod && startPeriod) endPeriod = startPeriod;
    if (endPeriod === "pm" && endHour < 12) endHour += 12;
    if (endPeriod === "am" && endHour === 12) endHour = 0;

    // eventDate is already the correct date from the calendar
    const startDate = new Date(eventDate);
    startDate.setHours(startHour, startMin, 0, 0);

    const endDate = new Date(eventDate);
    endDate.setHours(endHour, endMin, 0, 0);

    const summary = courseCode
      ? `${courseCode} - ${activityName}`
      : activityName;

    const dayNames = [
      "SUNDAY",
      "MONDAY",
      "TUESDAY",
      "WEDNESDAY",
      "THURSDAY",
      "FRIDAY",
      "SATURDAY",
    ];
    const dayOfWeek = dayNames[eventDate.getDay()];

    return {
      summary,
      startDate,
      endDate,
      location,
      description: `${courseCode}\n${activityName}`,
      recurrence: recurrence.toUpperCase() === "WEEKLY" ? "WEEKLY" : null,
      dayOfWeek,
    };
  }

  function filterEventsByDateRange(events, startDate, endDate) {
    return events.filter((event) => {
      if (event.recurrence === "WEEKLY") {
        return true;
      }

      const eventDate = event.startDate;
      return eventDate >= startDate && eventDate <= endDate;
    });
  }

  /**
   * Generate iCalendar file content
   */
  function generateICS(events, startDate, endDate) {
    const lines = [];

    lines.push("BEGIN:VCALENDAR");
    lines.push("VERSION:2.0");
    lines.push("PRODID:-//MQ Timetable Exporter//EN");
    lines.push("CALSCALE:GREGORIAN");
    lines.push("METHOD:PUBLISH");
    lines.push("X-WR-CALNAME:MQ Timetable");
    lines.push("X-WR-TIMEZONE:Australia/Sydney");

    lines.push("BEGIN:VTIMEZONE");
    lines.push("TZID:Australia/Sydney");
    lines.push("BEGIN:STANDARD");
    lines.push("DTSTART:19700101T000000");
    lines.push("TZOFFSETFROM:+1100");
    lines.push("TZOFFSETTO:+1000");
    lines.push("RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=1SU");
    lines.push("END:STANDARD");
    lines.push("BEGIN:DAYLIGHT");
    lines.push("DTSTART:19700101T000000");
    lines.push("TZOFFSETFROM:+1000");
    lines.push("TZOFFSETTO:+1100");
    lines.push("RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=1SU");
    lines.push("END:DAYLIGHT");
    lines.push("END:VTIMEZONE");

    for (const event of events) {
      lines.push("BEGIN:VEVENT");
      lines.push(`UID:${generateUID()}`);
      lines.push(`DTSTAMP:${formatDateTimeForICS(new Date())}`);
      lines.push(
        `DTSTART;TZID=Australia/Sydney:${formatDateTimeForICS(event.startDate)}`,
      );
      lines.push(
        `DTEND;TZID=Australia/Sydney:${formatDateTimeForICS(event.endDate)}`,
      );
      lines.push(`SUMMARY:${escapeICSText(event.summary)}`);

      if (event.location) {
        lines.push(`LOCATION:${escapeICSText(event.location)}`);
      }

      if (event.description) {
        lines.push(`DESCRIPTION:${escapeICSText(event.description)}`);
      }

      if (event.recurrence === "WEEKLY") {
        const until = formatDateTimeForICS(endDate);
        lines.push(`RRULE:FREQ=WEEKLY;UNTIL=${until}`);
      }

      lines.push("STATUS:CONFIRMED");
      lines.push("SEQUENCE:0");
      lines.push("END:VEVENT");
    }

    lines.push("END:VCALENDAR");

    return lines.join("\r\n");
  }

  /**
   * Download ICS file
   */
  function downloadICS(content, startDate, endDate) {
    const blob = new Blob([content], { type: "text/calendar;charset=utf-8" });
    const url = URL.createObjectURL(blob);

    const startStr = formatDateForFilename(startDate);
    const endStr = formatDateForFilename(endDate);
    const filename = `timetable_${startStr}_${endStr}.ics`;

    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.style.display = "none";

    document.body.appendChild(link);
    link.click();

    // Clean up
    setTimeout(() => {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 100);
  }

  // ===== UI CREATION =====

  function createExportButton() {
    const button = document.createElement("button");
    button.className = `${EXTENSION_PREFIX}-btn`;
    button.style.cssText = `
      position: relative;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 6px 16px;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
      font-size: 14px;
      font-weight: 500;
      color: #ffffff;
      background: linear-gradient(180deg, #0066cc 0%, #0052a3 100%);
      border: 1px solid #004a94;
      border-radius: 4px;
      cursor: pointer;
      transition: all 0.2s ease;
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);
      z-index: 1000;
    `;
    button.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" style="width: 16px; height: 16px; fill: currentColor;">
        <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
      </svg>
      Export
    `;

    button.addEventListener("click", togglePopover);
    button.addEventListener("mouseenter", function () {
      this.style.background =
        "linear-gradient(180deg, #0052a3 0%, #004080 100%)";
      this.style.boxShadow = "0 2px 4px rgba(0, 0, 0, 0.15)";
    });
    button.addEventListener("mouseleave", function () {
      this.style.background =
        "linear-gradient(180deg, #0066cc 0%, #0052a3 100%)";
      this.style.boxShadow = "0 1px 2px rgba(0, 0, 0, 0.1)";
    });

    return button;
  }

  function createPopover() {
    const popover = document.createElement("div");
    popover.className = `${EXTENSION_PREFIX}-popover`;
    popover.style.cssText = `
      position: absolute;
      top: calc(100% + 8px);
      right: 0;
      width: 340px;
      background: #ffffff;
      border: none;
      border-radius: 8px;
      box-shadow: 0 10px 40px rgba(0, 0, 0, 0.16), 0 0 0 1px rgba(0, 0, 0, 0.08);
      z-index: 10001;
      padding: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
      pointer-events: auto;
      overflow: hidden;
      animation: none;
      display: none;
    `;
    popover.addEventListener("click", (e) => {
      e.stopPropagation();
    });

    const today = new Date();
    const nextMonth = new Date(today);
    nextMonth.setMonth(nextMonth.getMonth() + 1);

    const todayStr = today.toISOString().split("T")[0];
    const nextMonthStr = nextMonth.toISOString().split("T")[0];

    popover.innerHTML = `
      <h3 style="margin: 0; padding: 16px 20px; font-size: 16px; font-weight: 600; color: #1a1a1a; border-bottom: 1px solid #f0f0f0; background: linear-gradient(135deg, #fafbfc 0%, #ffffff 100%);">Export Timetable</h3>
      <form class="${EXTENSION_PREFIX}-popover-form" id="${EXTENSION_PREFIX}-form" style="display: flex; flex-direction: column; gap: 14px; padding: 20px;">
        <div class="${EXTENSION_PREFIX}-form-group" style="display: flex; flex-direction: column; gap: 8px;">
          <label for="${EXTENSION_PREFIX}-start-date" style="font-size: 13px; font-weight: 600; color: #1a1a1a; text-transform: uppercase; letter-spacing: 0.3px;">Start Date</label>
          <input 
            type="date" 
            id="${EXTENSION_PREFIX}-start-date" 
            value="${todayStr}"
            required
            style="padding: 10px 12px; font-size: 14px; color: #1a1a1a; background: #ffffff; border: 1px solid #e0e0e0; border-radius: 6px; outline: none; transition: all 0.2s ease; font-family: inherit;"
          />
        </div>
        <div class="${EXTENSION_PREFIX}-form-group" style="display: flex; flex-direction: column; gap: 8px;">
          <label for="${EXTENSION_PREFIX}-end-date" style="font-size: 13px; font-weight: 600; color: #1a1a1a; text-transform: uppercase; letter-spacing: 0.3px;">End Date</label>
          <input 
            type="date" 
            id="${EXTENSION_PREFIX}-end-date" 
            value="${nextMonthStr}"
            required
            style="padding: 10px 12px; font-size: 14px; color: #1a1a1a; background: #ffffff; border: 1px solid #e0e0e0; border-radius: 6px; outline: none; transition: all 0.2s ease; font-family: inherit;"
          />
        </div>
        <div style="display: flex; gap: 12px; padding: 16px 20px; margin: -20px -20px 0 -20px; border-top: 1px solid #f0f0f0; background: #fafbfc;">
          <button 
            type="button" 
            class="${EXTENSION_PREFIX}-popover-btn ${EXTENSION_PREFIX}-popover-btn-secondary"
            id="${EXTENSION_PREFIX}-cancel-btn"
            style="flex: 1; padding: 10px 14px; font-size: 14px; font-weight: 500; border: 1px solid #e0e0e0; border-radius: 6px; cursor: pointer; transition: all 0.2s ease; outline: none; color: #333; background: #ffffff;"
          >
            Cancel
          </button>
          <button 
            type="submit" 
            class="${EXTENSION_PREFIX}-popover-btn ${EXTENSION_PREFIX}-popover-btn-primary"
            id="${EXTENSION_PREFIX}-export-btn"
            style="flex: 1; padding: 10px 14px; font-size: 14px; font-weight: 500; border: 1px solid #004a94; border-radius: 6px; cursor: pointer; transition: all 0.2s ease; outline: none; color: #ffffff; background: linear-gradient(180deg, #0066cc 0%, #0052a3 100%); box-shadow: 0 2px 4px rgba(0, 102, 204, 0.2);"
          >
            Export
          </button>
        </div>
        <div id="${EXTENSION_PREFIX}-message" style="display: none;"></div>
      </form>
    `;

    // Add event listeners
    const form = popover.querySelector(`#${EXTENSION_PREFIX}-form`);
    const cancelBtn = popover.querySelector(`#${EXTENSION_PREFIX}-cancel-btn`);

    if (!form || !cancelBtn) {
      return popover;
    }

    form.addEventListener("submit", handleExport);
    cancelBtn.addEventListener("click", closePopover);

    // Validate dates on change
    const startInput = popover.querySelector(`#${EXTENSION_PREFIX}-start-date`);
    const endInput = popover.querySelector(`#${EXTENSION_PREFIX}-end-date`);

    if (startInput) {
      startInput.addEventListener("change", validateDates);
      startInput.addEventListener("mouseenter", function () {
        this.style.borderColor = "#d0d0d0";
      });
      startInput.addEventListener("mouseleave", function () {
        this.style.borderColor = "#e0e0e0";
      });
      startInput.addEventListener("focus", function () {
        this.style.borderColor = "#0066cc";
        this.style.boxShadow = "0 0 0 3px rgba(0, 102, 204, 0.1)";
      });
      startInput.addEventListener("blur", function () {
        this.style.borderColor = "#e0e0e0";
        this.style.boxShadow = "none";
      });
    }
    if (endInput) {
      endInput.addEventListener("change", validateDates);
      endInput.addEventListener("mouseenter", function () {
        this.style.borderColor = "#d0d0d0";
      });
      endInput.addEventListener("mouseleave", function () {
        this.style.borderColor = "#e0e0e0";
      });
      endInput.addEventListener("focus", function () {
        this.style.borderColor = "#0066cc";
        this.style.boxShadow = "0 0 0 3px rgba(0, 102, 204, 0.1)";
      });
      endInput.addEventListener("blur", function () {
        this.style.borderColor = "#e0e0e0";
        this.style.boxShadow = "none";
      });
    }

    // Add hover states to buttons
    const cancelBtnEl = popover.querySelector(
      `#${EXTENSION_PREFIX}-cancel-btn`,
    );
    const exportBtnEl = popover.querySelector(
      `#${EXTENSION_PREFIX}-export-btn`,
    );

    if (cancelBtnEl) {
      cancelBtnEl.addEventListener("mouseenter", function () {
        this.style.background = "#f5f5f5";
        this.style.borderColor = "#d0d0d0";
      });
      cancelBtnEl.addEventListener("mouseleave", function () {
        this.style.background = "#ffffff";
        this.style.borderColor = "#e0e0e0";
      });
    }

    if (exportBtnEl) {
      exportBtnEl.addEventListener("mouseenter", function () {
        this.style.background =
          "linear-gradient(180deg, #0052a3 0%, #004080 100%)";
        this.style.boxShadow = "0 4px 8px rgba(0, 102, 204, 0.3)";
        this.style.transform = "translateY(-1px)";
      });
      exportBtnEl.addEventListener("mouseleave", function () {
        this.style.background =
          "linear-gradient(180deg, #0066cc 0%, #0052a3 100%)";
        this.style.boxShadow = "0 2px 4px rgba(0, 102, 204, 0.2)";
        this.style.transform = "translateY(0)";
      });
      exportBtnEl.addEventListener("mousedown", function () {
        this.style.transform = "translateY(0)";
        this.style.boxShadow = "0 2px 4px rgba(0, 102, 204, 0.2)";
      });
    }

    return popover;
  }

  function createOverlay() {
    const overlay = document.createElement("div");
    overlay.className = `${EXTENSION_PREFIX}-overlay`;
    overlay.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      z-index: 9999;
      background: transparent;
      pointer-events: auto;
      display: none;
    `;
    overlay.addEventListener("click", closePopover);
    return overlay;
  }

  function togglePopover(event) {
    event.stopPropagation();

    if (popoverOpen) {
      closePopover();
    } else {
      openPopover();
    }
  }

  function openPopover() {
    if (popover && overlay) {
      popover.style.display = "block";
      overlay.style.display = "block";
      popoverOpen = true;

      // Clear any previous messages
      const messageDiv = popover.querySelector(`#${EXTENSION_PREFIX}-message`);
      if (messageDiv) {
        messageDiv.style.display = "none";
        messageDiv.textContent = "";
        messageDiv.className = "";
      }
    }
  }

  function closePopover() {
    if (popover && overlay) {
      popover.style.display = "none";
      overlay.style.display = "none";
      popoverOpen = false;
    }
  }

  function validateDates() {
    const startInput = document.getElementById(
      `${EXTENSION_PREFIX}-start-date`,
    );
    const endInput = document.getElementById(`${EXTENSION_PREFIX}-end-date`);
    const exportBtn = document.getElementById(`${EXTENSION_PREFIX}-export-btn`);

    if (startInput && endInput && exportBtn) {
      const startDate = new Date(startInput.value);
      const endDate = new Date(endInput.value);

      exportBtn.disabled = endDate < startDate;
    }
  }

  function showMessage(message, type = "error") {
    const messageDiv = document.getElementById(`${EXTENSION_PREFIX}-message`);
    if (messageDiv) {
      messageDiv.textContent = message;
      let styles =
        "padding: 12px; margin-top: 10px; font-size: 13px; border-radius: 6px; line-height: 1.5;";

      if (type === "success") {
        messageDiv.className = `${EXTENSION_PREFIX}-success-message`;
        styles +=
          " color: #2f855a; background: #f0fff4; border: 1px solid #9ae6b4;";
      } else if (type === "info") {
        messageDiv.className = `${EXTENSION_PREFIX}-info-message`;
        styles +=
          " color: #2c5282; background: #ebf8ff; border: 1px solid #90cdf4;";
      } else {
        messageDiv.className = `${EXTENSION_PREFIX}-error-message`;
        styles +=
          " color: #c53030; background: #fff5f5; border: 1px solid #feb2b2;";
      }

      messageDiv.style.cssText = styles;
      messageDiv.style.display = "block";
    }
  }

  async function handleExport(event) {
    event.preventDefault();

    const startInput = document.getElementById(
      `${EXTENSION_PREFIX}-start-date`,
    );
    const endInput = document.getElementById(`${EXTENSION_PREFIX}-end-date`);
    const exportBtn = document.getElementById(`${EXTENSION_PREFIX}-export-btn`);

    if (!startInput || !endInput || !exportBtn) {
      return;
    }

    const startDate = new Date(startInput.value);
    startDate.setHours(0, 0, 0, 0);

    const endDate = new Date(endInput.value);
    endDate.setHours(23, 59, 59, 999);

    if (endDate < startDate) {
      showMessage("End date must be after start date", "error");
      return;
    }

    exportBtn.disabled = true;
    exportBtn.innerHTML =
      '<span class="mq-loading-spinner"></span> Exporting...';

    try {
      const allEvents = extractTimetableData(startDate);

      if (allEvents.length === 0) {
        showMessage("No timetable data found on this page", "error");
        return;
      }

      showMessage(
        `Found ${allEvents.length} event${allEvents.length === 1 ? "" : "s"}`,
        "info",
      );

      const filteredEvents = filterEventsByDateRange(
        allEvents,
        startDate,
        endDate,
      );

      if (filteredEvents.length === 0) {
        showMessage(
          `No classes found between ${startInput.value} and ${endInput.value}`,
          "error",
        );
        return;
      }

      const icsContent = generateICS(filteredEvents, startDate, endDate);

      downloadICS(icsContent, startDate, endDate);

      showMessage(
        `Successfully exported ${filteredEvents.length} class(es)`,
        "success",
      );

      setTimeout(() => {
        closePopover();
      }, 1500);
    } catch (error) {
      showMessage("Failed to export timetable. Please try again.", "error");
    } finally {
      exportBtn.disabled = false;
      exportBtn.textContent = "Export";
    }
  }

  function findInjectionPoint() {
    const bannerRight = document.querySelector(".bannerRight");
    if (bannerRight) {
      return { element: bannerRight, position: "prepend" };
    }

    const selectors = [
      "header .actions",
      "header .toolbar",
      ".page-header .actions",
      ".page-header .toolbar",
      '[class*="header"] [class*="action"]',
      '[class*="toolbar"]',
      "header",
      ".header",
    ];

    for (const selector of selectors) {
      const element = document.querySelector(selector);
      if (element) {
        return { element, position: "append" };
      }
    }

    // Fallback: inject at the top of body
    return { element: document.body, position: "fixed" };
  }

  function injectUI() {
    // Check if already injected
    if (document.querySelector(`.${EXTENSION_PREFIX}-btn`)) {
      return;
    }

    // Create elements
    exportButton = createExportButton();
    popover = createPopover();
    overlay = createOverlay();

    const injectionPoint = findInjectionPoint();

    if (injectionPoint.position === "fixed") {
      // Create a container for absolute positioning
      const container = document.createElement("div");
      container.style.position = "fixed";
      container.style.top = "20px";
      container.style.right = "20px";
      container.style.zIndex = "10000";
      container.style.pointerEvents = "auto";

      container.appendChild(exportButton);
      exportButton.appendChild(popover);

      document.body.appendChild(overlay);
      document.body.appendChild(container);
    } else if (injectionPoint.position === "prepend") {
      // Inject at the beginning of the element
      const container = document.createElement("div");
      container.style.display = "inline-flex";
      container.style.alignItems = "center";
      container.style.marginRight = "15px";
      container.style.position = "relative";
      container.style.zIndex = "10000";
      container.style.pointerEvents = "auto";

      container.appendChild(exportButton);
      exportButton.appendChild(popover);

      injectionPoint.element.insertBefore(
        container,
        injectionPoint.element.firstChild,
      );
      document.body.appendChild(overlay);
    } else {
      // Inject relative to the found element (append)
      const container = document.createElement("div");
      container.style.position = "relative";
      container.style.display = "inline-block";
      container.style.marginLeft = "10px";
      container.style.zIndex = "10000";
      container.style.pointerEvents = "auto";

      container.appendChild(exportButton);
      exportButton.appendChild(popover);

      injectionPoint.element.appendChild(container);
      document.body.appendChild(overlay);
    }
  }

  /**
   * Initialize the extension
   */
  function init() {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", init);
      return;
    }

    setTimeout(() => {
      injectUI();
    }, 1000);
  }

  init();
})();
