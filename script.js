const header = document.querySelector(".site-header");
const navToggle = document.querySelector(".nav-toggle");
const navLinks = document.querySelectorAll(".site-nav a");
const year = document.querySelector("#year");
const CHATBASE_DEFAULT_BOT_ID = "wDye2cLkm_hH2e4lyjq_F";
const CHATBASE_DEFAULT_HOST = "www.chatbase.co";
const themeStorageKey = "bansal-stratedge-theme";
const root = document.documentElement;
const themeMediaQuery =
  typeof window.matchMedia === "function"
    ? window.matchMedia("(prefers-color-scheme: dark)")
    : null;

const readStoredTheme = () => {
  try {
    const storedTheme = window.localStorage.getItem(themeStorageKey);

    if (storedTheme === "dark" || storedTheme === "light") {
      return storedTheme;
    }
  } catch (error) {
    return null;
  }

  return null;
};

const getSystemTheme = () => (themeMediaQuery?.matches ? "dark" : "light");

const getActiveTheme = () => {
  const storedTheme = readStoredTheme();

  if (storedTheme) {
    return storedTheme;
  }

  if (root.dataset.theme === "dark" || root.dataset.theme === "light") {
    return root.dataset.theme;
  }

  return getSystemTheme();
};

const syncThemeToggle = (theme) => {
  document.querySelectorAll(".theme-toggle").forEach((toggle) => {
    const label = toggle.querySelector("[data-theme-label]");
    const nextTheme = theme === "dark" ? "light" : "dark";

    if (label) {
      label.textContent = theme === "dark" ? "Dark mode" : "Light mode";
    }

    toggle.setAttribute("aria-label", `Switch to ${nextTheme} mode`);
    toggle.setAttribute("aria-pressed", String(theme === "dark"));
    toggle.setAttribute("title", `Switch to ${nextTheme} mode`);
  });
};

const applyTheme = (theme, { persist = false } = {}) => {
  root.dataset.theme = theme;
  root.style.colorScheme = theme;

  if (persist) {
    try {
      window.localStorage.setItem(themeStorageKey, theme);
    } catch (error) {
      // Ignore localStorage write issues and still apply the theme for this session.
    }
  }

  syncThemeToggle(theme);
};

const createThemeToggle = () => {
  const navShell = document.querySelector(".nav-shell");
  const siteNav = navShell?.querySelector(".site-nav");

  if (!navShell || navShell.querySelector(".theme-toggle")) {
    return;
  }

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "theme-toggle";
  toggle.innerHTML = `
    <span class="theme-toggle-track" aria-hidden="true">
      <span class="theme-toggle-stars">
        <span class="theme-toggle-star theme-toggle-star-a"></span>
        <span class="theme-toggle-star theme-toggle-star-b"></span>
        <span class="theme-toggle-star theme-toggle-star-c"></span>
        <span class="theme-toggle-star theme-toggle-star-d"></span>
        <span class="theme-toggle-star theme-toggle-star-e"></span>
      </span>
      <span class="theme-toggle-cloud theme-toggle-cloud-one"></span>
      <span class="theme-toggle-cloud theme-toggle-cloud-two"></span>
      <span class="theme-toggle-horizon"></span>
      <span class="theme-toggle-thumb">
        <span class="theme-toggle-crater theme-toggle-crater-one"></span>
        <span class="theme-toggle-crater theme-toggle-crater-two"></span>
        <span class="theme-toggle-crater theme-toggle-crater-three"></span>
      </span>
    </span>
    <span class="theme-toggle-label" data-theme-label></span>
  `;

  toggle.addEventListener("click", () => {
    const nextTheme = getActiveTheme() === "dark" ? "light" : "dark";
    applyTheme(nextTheme, { persist: true });
  });

  if (siteNav) {
    siteNav.insertAdjacentElement("afterend", toggle);
  } else {
    navShell.appendChild(toggle);
  }

  syncThemeToggle(getActiveTheme());
};

const initTheme = () => {
  applyTheme(getActiveTheme());
  createThemeToggle();

  if (!themeMediaQuery) {
    return;
  }

  const handleThemeChange = (event) => {
    if (readStoredTheme()) {
      return;
    }

    applyTheme(event.matches ? "dark" : "light");
  };

  if (typeof themeMediaQuery.addEventListener === "function") {
    themeMediaQuery.addEventListener("change", handleThemeChange);
  } else if (typeof themeMediaQuery.addListener === "function") {
    themeMediaQuery.addListener(handleThemeChange);
  }
};

const getChatbaseConfig = () => {
  const botIdMeta = document.querySelector('meta[name="chatbase-bot-id"]');
  const hostMeta = document.querySelector('meta[name="chatbase-host"]');
  const runtimeConfig = window.CHATBASE_CONFIG || {};

  return {
    botId:
      window.CHATBASE_BOT_ID ||
      runtimeConfig.botId ||
      botIdMeta?.content ||
      CHATBASE_DEFAULT_BOT_ID,
    host:
      window.CHATBASE_HOST ||
      runtimeConfig.host ||
      hostMeta?.content ||
      CHATBASE_DEFAULT_HOST,
  };
};

const ensureChatbaseLoaded = () => {
  const { botId, host } = getChatbaseConfig();

  if (!botId || document.getElementById(botId)) {
    return;
  }

  const existingState =
    typeof window.chatbase === "function" ? window.chatbase("getState") : null;

  if (existingState === "initialized") {
    return;
  }

  if (!window.chatbase) {
    window.chatbase = (...args) => {
      window.chatbase.q = window.chatbase.q || [];
      window.chatbase.q.push(args);
    };

    window.chatbase = new Proxy(window.chatbase, {
      get(target, prop) {
        if (prop === "q") {
          return target.q;
        }

        return (...args) => target(prop, ...args);
      },
    });
  }

  const script = document.createElement("script");
  script.src = `https://${host}/embed.min.js`;
  script.id = botId;
  script.setAttribute("domain", host);
  document.body.appendChild(script);
};

if (year) {
  year.textContent = new Date().getFullYear();
}

initTheme();

if (document.readyState === "complete") {
  ensureChatbaseLoaded();
} else {
  window.addEventListener("load", ensureChatbaseLoaded, { once: true });
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

const normalizePagePath = (pathname) => pathname.split("/").pop() || "index.html";
const getNavLinkTarget = (link) => {
  const href = link.getAttribute("href");

  if (!href) {
    return { page: "", hash: "" };
  }

  const targetUrl = new URL(href, window.location.href);

  return {
    page: normalizePagePath(targetUrl.pathname),
    hash: targetUrl.hash,
  };
};
const setCurrentNavLink = (matcher) => {
  let hasMatch = false;

  navLinks.forEach((link) => {
    const isMatch = !hasMatch && matcher(link);

    if (isMatch) {
      link.setAttribute("aria-current", "page");
      hasMatch = true;
      return;
    }

    link.removeAttribute("aria-current");
  });
};
const path = normalizePagePath(window.location.pathname);
const isHomePage = document.body.classList.contains("home-page");
const isFoundersPage = document.body.classList.contains("founders-page");
const isWorkshopsPage = document.body.classList.contains("workshops-page");
const isBlogsSection =
  document.body.classList.contains("blogs-page") ||
  document.body.classList.contains("journal-article-page") ||
  path === "blogs.html" ||
  path.startsWith("journal-");
const updateCurrentNavLink = () => {
  const currentHash = window.location.hash;
  const isCurrentHomePage = isHomePage || path === "index.html";

  if (isCurrentHomePage && (currentHash === "#capabilities" || currentHash === "#approach" || currentHash === "#contact")) {
    setCurrentNavLink((link) => {
      const { page, hash } = getNavLinkTarget(link);
      return page === "index.html" && hash === currentHash;
    });
    return;
  }

  if (isBlogsSection) {
    setCurrentNavLink((link) => {
      const { page, hash } = getNavLinkTarget(link);
      return page === "blogs.html" && hash === "";
    });
    return;
  }

  if (isFoundersPage || path === "founders.html") {
    setCurrentNavLink((link) => {
      const { page, hash } = getNavLinkTarget(link);
      return page === "founders.html" && hash === "";
    });
    return;
  }

  if (isWorkshopsPage || path === "workshops.html") {
    setCurrentNavLink((link) => {
      const { page, hash } = getNavLinkTarget(link);
      return page === "workshops.html" && hash === "";
    });
    return;
  }

  if (isCurrentHomePage) {
    setCurrentNavLink((link) => {
      const { page, hash } = getNavLinkTarget(link);
      return page === "index.html" && hash === "";
    });
    return;
  }

  setCurrentNavLink(() => false);
};

updateCurrentNavLink();
window.addEventListener("hashchange", updateCurrentNavLink);

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
const workshopInterestForm = document.querySelector("#workshop-interest-form");
const workshopName = document.querySelector("#workshop-name");
const workshopEmail = document.querySelector("#workshop-email");
const workshopOrganization = document.querySelector("#workshop-organization");
const workshopFormat = document.querySelector("#workshop-format");
const workshopAudience = document.querySelector("#workshop-audience");
const workshopDelivery = document.querySelector("#workshop-delivery");
const workshopNote = document.querySelector("#workshop-note");
const workshopResponse = document.querySelector("#workshop-response");

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

if (workshopInterestForm && workshopName && workshopEmail && workshopResponse) {
  workshopInterestForm.addEventListener("submit", (event) => {
    event.preventDefault();

    if (!workshopInterestForm.reportValidity()) {
      return;
    }

    const name = workshopName.value.trim();
    const email = workshopEmail.value.trim();
    const organization = workshopOrganization ? workshopOrganization.value.trim() : "";
    const format = workshopFormat ? workshopFormat.value.trim() : "";
    const audience = workshopAudience ? workshopAudience.value.trim() : "";
    const delivery = workshopDelivery ? workshopDelivery.value.trim() : "";
    const note = workshopNote ? workshopNote.value.trim() : "";
    const subject = organization ? `Workshop or Speaking Inquiry - ${organization}` : "Workshop or Speaking Inquiry";
    const body =
      `Hi Kamlesh,\n\nI'm interested in your workshop or speaking offerings.\n\nName: ${name}\nEmail: ${email}\n${
        organization ? `Institution / company: ${organization}\n` : ""
      }${format ? `Preferred format: ${format}\n` : ""}${audience ? `Audience: ${audience}\n` : ""}${
        delivery ? `Delivery mode: ${delivery}\n` : ""
      }${note ? `\nInterest / timing / pricing note:\n${note}\n` : ""}\nPlease share availability and indicative pricing.\n\nThanks.`;

    workshopResponse.textContent =
      "Gmail is opening with a prefilled workshop or speaking inquiry to Kamlesh.";

    openGmailCompose({
      to: "kamlesh@kamleshbansal.com",
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
