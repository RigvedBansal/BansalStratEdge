const header = document.querySelector(".site-header");
const navToggle = document.querySelector(".nav-toggle");
const navLinks = document.querySelectorAll(".site-nav a");
const year = document.querySelector("#year");

if (year) {
  year.textContent = new Date().getFullYear();
}

if (navToggle && header) {
  navToggle.addEventListener("click", () => {
    const expanded = navToggle.getAttribute("aria-expanded") === "true";
    navToggle.setAttribute("aria-expanded", String(!expanded));
    header.dataset.navOpen = String(!expanded);
  });
}

navLinks.forEach((link) => {
  link.addEventListener("click", () => {
    if (!header || !navToggle) {
      return;
    }

    header.dataset.navOpen = "false";
    navToggle.setAttribute("aria-expanded", "false");
  });
});

const path = window.location.pathname.split("/").pop() || "index.html";
const isHomePage = document.body.classList.contains("home-page");
const isFoundersPage = document.body.classList.contains("founders-page");
const isBlogsSection =
  document.body.classList.contains("blogs-page") ||
  document.body.classList.contains("journal-article-page") ||
  path === "blogs.html" ||
  path.startsWith("journal-");

document.querySelectorAll("[data-nav-link]").forEach((link) => {
  const target = link.getAttribute("href");

  if ((isHomePage || path === "index.html" || path === "") && target === "./index.html") {
    link.setAttribute("aria-current", "page");
  }

  if (isBlogsSection && target === "./blogs.html") {
    link.setAttribute("aria-current", "page");
  }

  if ((isFoundersPage || path === "founders.html") && target === "./founders.html") {
    link.setAttribute("aria-current", "page");
  }
});

const reveals = document.querySelectorAll(".reveal");

if ("IntersectionObserver" in window) {
  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) {
          return;
        }

        entry.target.classList.add("is-visible");
        observer.unobserve(entry.target);
      });
    },
    {
      threshold: 0.14,
      rootMargin: "0px 0px -8% 0px",
    }
  );

  reveals.forEach((item) => observer.observe(item));
} else {
  reveals.forEach((item) => item.classList.add("is-visible"));
}

const checklistForm = document.querySelector("#checklist-form");
const checklistEmail = document.querySelector("#checklist-email");
const checklistResponse = document.querySelector("#checklist-response");
const strategyForm = document.querySelector("#strategy-call-form");
const strategyDate = document.querySelector("#strategy-date");
const strategyTime = document.querySelector("#strategy-time");
const strategyCompany = document.querySelector("#strategy-company");
const strategyNote = document.querySelector("#strategy-note");
const strategyResponse = document.querySelector("#strategy-response");
const callPlanner = document.querySelector("#call-planner");
const checklistToggles = document.querySelectorAll(".checklist-toggle");
const checklistSections = document.querySelectorAll("[data-checklist-section]");
const checklistCompleteCount = document.querySelector("#checklist-complete-count");
const checklistTotalCount = document.querySelector("#checklist-total-count");
const checklistProgressBar = document.querySelector("#checklist-progress-bar");
const checklistStatusCopy = document.querySelector("#checklist-status-copy");
const checklistResetButton = document.querySelector("[data-checklist-reset]");
const checklistStorageKey = "bansal-stratedge-cfo-checklist-progress";

const formatDateLabel = (dateValue) => {
  if (!dateValue) {
    return "";
  }

  const parsedDate = new Date(`${dateValue}T00:00:00`);

  if (Number.isNaN(parsedDate.getTime())) {
    return dateValue;
  }

  return parsedDate.toLocaleDateString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric",
  });
};

const formatTimeLabel = (timeValue) => {
  if (!timeValue) {
    return "";
  }

  const [hours, minutes] = timeValue.split(":").map(Number);

  if (Number.isNaN(hours) || Number.isNaN(minutes)) {
    return timeValue;
  }

  const parsedTime = new Date();
  parsedTime.setHours(hours, minutes, 0, 0);

  return parsedTime.toLocaleTimeString(undefined, {
    hour: "numeric",
    minute: "2-digit",
  });
};

const buildGmailComposeUrl = ({ to, subject, body }) => {
  const params = new URLSearchParams({
    view: "cm",
    fs: "1",
    to,
    su: subject,
    body,
  });

  return `https://mail.google.com/mail/?${params.toString()}`;
};

const openGmailCompose = ({ to, subject, body }) => {
  window.location.assign(buildGmailComposeUrl({ to, subject, body }));
};

if (strategyDate) {
  const today = new Date().toISOString().split("T")[0];
  strategyDate.min = today;
}

document.querySelectorAll("[data-scroll-to-call]").forEach((link) => {
  link.addEventListener("click", (event) => {
    if (!callPlanner) {
      return;
    }

    event.preventDefault();

    callPlanner.scrollIntoView({
      behavior: "smooth",
      block: "start",
    });

    window.history.replaceState(null, "", "#call-planner");

    if (strategyDate) {
      window.setTimeout(() => strategyDate.focus(), 450);
    }
  });
});

if (strategyForm && strategyDate && strategyTime && strategyResponse) {
  strategyForm.addEventListener("submit", (event) => {
    event.preventDefault();

    if (!strategyForm.reportValidity()) {
      return;
    }

    const selectedDate = formatDateLabel(strategyDate.value);
    const selectedTime = formatTimeLabel(strategyTime.value);
    const company = strategyCompany ? strategyCompany.value.trim() : "";
    const note = strategyNote ? strategyNote.value.trim() : "";
    const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone || "Local timezone";
    const subject = company ? `Strategy Call Request - ${company}` : "Strategy Call Request";
    const body =
      `Hi Kamlesh,\n\nI'd like to book a strategy call with Bansal StratEdge.\n\nPreferred date: ${selectedDate}\nPreferred time: ${selectedTime}\nTimezone: ${timezone}\n${
        company ? `Company / role: ${company}\n` : ""
      }${note ? `\nWhat I want to discuss:\n${note}\n` : ""}\nThanks.`;

    strategyResponse.textContent =
      "Gmail is opening with a prefilled draft to Kamlesh. If Gmail is not signed in yet, log in and the draft should still be ready.";

    openGmailCompose({
      to: "kamlesh@kamleshbansal.com",
      subject,
      body,
    });
  });
}

if (checklistForm && checklistEmail && checklistResponse) {
  checklistForm.addEventListener("submit", (event) => {
    event.preventDefault();

    if (!checklistEmail.reportValidity()) {
      return;
    }

    const email = checklistEmail.value.trim();
    const subject = "Checklist Request - CFO Capital Efficiency Checklist";
    const body = `Hi Rigved,\n\nPlease send the CFO Capital Efficiency Checklist to this requester.\n\nRequester email: ${email}\n\nThanks.`;

    checklistResponse.textContent =
      "Gmail is opening with a prefilled request to rigvedkbansal@gmail.com.";

    openGmailCompose({
      to: "rigvedkbansal@gmail.com",
      subject,
      body,
    });
  });
}

const getChecklistStatusMessage = (checkedCount, totalCount) => {
  if (!totalCount || checkedCount === 0) {
    return "Start with the checks that would change a decision this month. Progress saves on this device.";
  }

  const ratio = checkedCount / totalCount;

  if (ratio < 0.35) {
    return "The biggest gains are probably still in cash visibility, approval discipline, and reporting clarity.";
  }

  if (ratio < 0.7) {
    return "The base is forming. Tighten ownership, cadence, and stop-or-reset reviews to remove the remaining drag.";
  }

  if (checkedCount < totalCount) {
    return "The operating base looks stronger. The next gains usually come from sharper automation and capital prioritization.";
  }

  return "Strong pass. The remaining work is less about basics and more about compounding discipline.";
};

const readChecklistState = () => {
  try {
    return JSON.parse(window.localStorage.getItem(checklistStorageKey) || "{}");
  } catch (error) {
    return {};
  }
};

const writeChecklistState = (state) => {
  try {
    window.localStorage.setItem(checklistStorageKey, JSON.stringify(state));
  } catch (error) {
    // Ignore storage failures so the checklist still works as a plain page.
  }
};

const updateChecklistProgress = () => {
  if (!checklistToggles.length) {
    return;
  }

  const totalCount = checklistToggles.length;
  const checkedCount = Array.from(checklistToggles).filter((toggle) => toggle.checked).length;
  const completionPercent = Math.round((checkedCount / totalCount) * 100);

  if (checklistCompleteCount) {
    checklistCompleteCount.textContent = String(checkedCount);
  }

  if (checklistTotalCount) {
    checklistTotalCount.textContent = String(totalCount);
  }

  if (checklistProgressBar) {
    checklistProgressBar.style.width = `${completionPercent}%`;
  }

  if (checklistStatusCopy) {
    checklistStatusCopy.textContent = getChecklistStatusMessage(checkedCount, totalCount);
  }

  checklistSections.forEach((section) => {
    const sectionToggles = section.querySelectorAll(".checklist-toggle");
    const sectionCheckedCount = Array.from(sectionToggles).filter((toggle) => toggle.checked).length;
    const sectionStatus = section.querySelector("[data-section-status]");

    section.classList.toggle("is-complete", sectionCheckedCount === sectionToggles.length && sectionToggles.length > 0);

    if (sectionStatus) {
      sectionStatus.textContent = `${sectionCheckedCount} / ${sectionToggles.length} done`;
    }
  });
};

if (checklistToggles.length) {
  const savedChecklistState = readChecklistState();

  checklistToggles.forEach((toggle) => {
    if (toggle.id && typeof savedChecklistState[toggle.id] === "boolean") {
      toggle.checked = savedChecklistState[toggle.id];
    }

    toggle.addEventListener("change", () => {
      const nextState = readChecklistState();
      nextState[toggle.id] = toggle.checked;
      writeChecklistState(nextState);
      updateChecklistProgress();
    });
  });

  if (checklistResetButton) {
    checklistResetButton.addEventListener("click", () => {
      const shouldReset = window.confirm("Reset all checklist selections?");

      if (!shouldReset) {
        return;
      }

      checklistToggles.forEach((toggle) => {
        toggle.checked = false;
      });

      writeChecklistState({});
      updateChecklistProgress();
    });
  }

  updateChecklistProgress();
}
