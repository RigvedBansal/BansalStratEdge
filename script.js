const header = document.querySelector(".site-header");
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

const ensureNavToggle = () => {
  const navShell = document.querySelector(".nav-shell");
  const siteNav = navShell?.querySelector(".site-nav");

  if (!navShell || !siteNav) {
    return;
  }

  if (!siteNav.id) {
    siteNav.id = "site-nav";
  }

  if (navShell.querySelector(".nav-toggle")) {
    return;
  }

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "nav-toggle";
  toggle.setAttribute("aria-expanded", "false");
  toggle.setAttribute("aria-controls", siteNav.id);
  toggle.setAttribute("aria-label", "Toggle navigation");
  toggle.innerHTML = "<span></span><span></span>";
  navShell.insertBefore(toggle, siteNav);
};

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

ensureNavToggle();
initTheme();

if (document.readyState === "complete") {
  ensureChatbaseLoaded();
} else {
  window.addEventListener("load", ensureChatbaseLoaded, { once: true });
}

const bindNavToggle = () => {
  const navToggle = document.querySelector(".nav-toggle");

  if (!navToggle || !header || navToggle.dataset.bound === "true") {
    return;
  }

  navToggle.dataset.bound = "true";
  navToggle.addEventListener("click", () => {
    const expanded = navToggle.getAttribute("aria-expanded") === "true";
    navToggle.setAttribute("aria-expanded", String(!expanded));
    header.dataset.navOpen = String(!expanded);
  });
};

bindNavToggle();

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
const isToolkitPage =
  document.body.classList.contains("toolkit-page") ||
  document.body.classList.contains("tool-detail-page") ||
  path === "finance-systems-toolkit.html" ||
  path === "resources-toolkit.html";
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

  if (isToolkitPage) {
    setCurrentNavLink((link) => {
      const { page, hash } = getNavLinkTarget(link);
      return page === "finance-systems-toolkit.html" && hash === "";
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
const copyButtons = document.querySelectorAll("[data-copy-target]");
const toolkitScoreInputs = document.querySelectorAll("[data-toolkit-score-input]");
const toolkitScoreGroups = document.querySelectorAll("[data-toolkit-score-group]");
const toolkitOverallScore = document.querySelector("#toolkit-overall-score");
const toolkitOverallFlag = document.querySelector("#toolkit-overall-flag");
const toolkitOverallProgressBar = document.querySelector("#toolkit-overall-progress-bar");
const toolkitExecutiveSummary = document.querySelector("#toolkit-executive-summary");
const toolkitPriorityList = document.querySelector("#toolkit-priority-list");
const toolkitResetButton = document.querySelector("[data-toolkit-reset]");
const toolkitStorageKey = "bansal-stratedge-toolkit-scorecard";

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

const copyTextToClipboard = async (text) => {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return true;
  }

  const tempField = document.createElement("textarea");
  tempField.value = text;
  tempField.setAttribute("readonly", "");
  tempField.style.position = "absolute";
  tempField.style.left = "-9999px";
  document.body.appendChild(tempField);
  tempField.select();

  const isCopied = document.execCommand("copy");
  document.body.removeChild(tempField);
  return isCopied;
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

if (copyButtons.length) {
  copyButtons.forEach((button) => {
    const originalLabel = button.textContent;

    button.addEventListener("click", async () => {
      const targetId = button.dataset.copyTarget;
      const target = targetId ? document.getElementById(targetId) : null;
      const textToCopy = target ? (target.value || target.textContent || "").trim() : "";

      if (!textToCopy) {
        return;
      }

      try {
        const isCopied = await copyTextToClipboard(textToCopy);
        button.textContent = isCopied ? "Copied" : "Copy failed";
      } catch (error) {
        button.textContent = "Copy failed";
      }

      window.setTimeout(() => {
        button.textContent = originalLabel;
      }, 1800);
    });
  });
}

const toolkitFlagClassNames = ["is-red", "is-amber", "is-green"];
const getToolkitFlag = (score) => {
  if (score < 2.5) {
    return { label: "Red", className: "is-red" };
  }

  if (score < 3.75) {
    return { label: "Amber", className: "is-amber" };
  }

  return { label: "Green", className: "is-green" };
};

const applyToolkitFlag = (element, score) => {
  if (!element) {
    return;
  }

  const { label, className } = getToolkitFlag(score);
  element.textContent = label;
  element.classList.add("toolkit-flag");
  element.classList.remove(...toolkitFlagClassNames);
  element.classList.add(className);
};

const readToolkitState = () => {
  try {
    return JSON.parse(window.localStorage.getItem(toolkitStorageKey) || "{}");
  } catch (error) {
    return {};
  }
};

const writeToolkitState = (state) => {
  try {
    window.localStorage.setItem(toolkitStorageKey, JSON.stringify(state));
  } catch (error) {
    // Ignore storage issues and keep the scorecard working in-memory.
  }
};

const getToolkitBoardImplication = (flagLabel, weakestSectionLabel) => {
  if (flagLabel === "Red") {
    return `Capital discipline is weak enough to slow decisions and raise avoidable risk. Start by stabilizing ${weakestSectionLabel}.`;
  }

  if (flagLabel === "Amber") {
    return "The finance base is workable, but ownership and cadence are uneven enough to keep creating drag.";
  }

  return "The operating base is strong enough to shift attention toward sharper optimization, automation, and capital prioritization.";
};

const getToolkitActionForGroup = (groupKey) => {
  const actions = {
    cash: "tighten the weekly 13-week cash cadence and align leadership on one liquidity view.",
    "working-capital": "assign owners for overdue receivables, renegotiate priority terms, and quantify trapped cash.",
    "spend-control": "separate committed and discretionary spend, refresh renewal reviews, and reset approval thresholds.",
    "capital-allocation": "run stop / continue / reset reviews on underperforming projects and restack the quarter's capital priorities.",
    reporting: "compress management reporting around driver movement, decision points, and next actions.",
    "ai-controls": "pilot only one repeatable AI workflow with named reviewer, prompt versioning, and an audit trail.",
  };

  return actions[groupKey] || "reset ownership, review cadence, and decision rules around this section.";
};

const buildToolkitExecutiveSummary = (overallAverage, orderedSections) => {
  const weakestSections = orderedSections.slice(0, 2);
  const strongestSections = orderedSections.slice(-2).reverse();
  const overallFlag = getToolkitFlag(overallAverage).label;
  const weakestLabel = weakestSections[0]?.label || "the weakest section";
  const priorityList = weakestSections.length
    ? weakestSections.map((section) => section.label).join("; ")
    : "None selected yet";
  const strengthList = strongestSections.length
    ? strongestSections.map((section) => section.label).join("; ")
    : "Still being rated";

  return [
    `Overall capital efficiency score: ${overallAverage.toFixed(1)} / 5 (${overallFlag}).`,
    `Priority sections: ${priorityList}.`,
    `Strongest sections: ${strengthList}.`,
    `Board-level implication: ${getToolkitBoardImplication(overallFlag, weakestLabel)}`,
    `Next 30 days: First, ${getToolkitActionForGroup(weakestSections[0]?.key)} Then, ${getToolkitActionForGroup(weakestSections[1]?.key)}`,
  ].join("\n");
};

const updateToolkitScorecard = () => {
  if (!toolkitScoreInputs.length) {
    return;
  }

  const sectionSummaries = [];

  toolkitScoreGroups.forEach((group, groupIndex) => {
    const groupInputs = group.querySelectorAll("[data-toolkit-score-input]");
    const groupLabel = group.querySelector(".suite-label")?.textContent?.trim() || `Section ${groupIndex + 1}`;
    const groupKey = groupInputs[0]?.dataset.toolkitGroup || `group-${groupIndex + 1}`;
    const groupTotal = Array.from(groupInputs).reduce((sum, input, inputIndex) => {
      const numericValue = Number(input.value) || 0;
      const scoreValue = input.parentElement?.querySelector("[data-toolkit-score-value]");

      if (scoreValue) {
        scoreValue.textContent = String(numericValue);
      }

      return sum + numericValue;
    }, 0);

    const groupAverage = groupInputs.length ? groupTotal / groupInputs.length : 0;
    const groupScore = group.querySelector("[data-toolkit-group-score]");
    const groupFlag = group.querySelector("[data-toolkit-group-flag]");

    if (groupScore) {
      groupScore.textContent = `${groupAverage.toFixed(1)} / 5`;
    }

    applyToolkitFlag(groupFlag, groupAverage);

    sectionSummaries.push({
      key: groupKey,
      label: groupLabel,
      average: groupAverage,
    });
  });

  const orderedSections = [...sectionSummaries].sort((left, right) => left.average - right.average);
  const overallAverage = sectionSummaries.length
    ? sectionSummaries.reduce((sum, section) => sum + section.average, 0) / sectionSummaries.length
    : 0;

  if (toolkitOverallScore) {
    toolkitOverallScore.textContent = overallAverage.toFixed(1);
  }

  applyToolkitFlag(toolkitOverallFlag, overallAverage);

  if (toolkitOverallProgressBar) {
    toolkitOverallProgressBar.style.width = `${Math.round((overallAverage / 5) * 100)}%`;
  }

  if (toolkitPriorityList) {
    toolkitPriorityList.innerHTML = orderedSections
      .slice(0, 3)
      .map((section) => `<li>${section.label}</li>`)
      .join("");
  }

  if (toolkitExecutiveSummary) {
    toolkitExecutiveSummary.value = buildToolkitExecutiveSummary(overallAverage, orderedSections);
  }
};

if (toolkitScoreInputs.length) {
  const savedToolkitState = readToolkitState();

  toolkitScoreInputs.forEach((input, inputIndex) => {
    const groupKey = input.dataset.toolkitGroup || "group";
    const storageId = `${groupKey}-${inputIndex}`;
    input.dataset.toolkitStorageId = storageId;

    if (savedToolkitState[storageId]) {
      input.value = String(savedToolkitState[storageId]);
    }

    input.addEventListener("input", () => {
      const nextState = readToolkitState();
      nextState[storageId] = Number(input.value) || 3;
      writeToolkitState(nextState);
      updateToolkitScorecard();
    });
  });

  if (toolkitResetButton) {
    toolkitResetButton.addEventListener("click", () => {
      const shouldReset = window.confirm("Reset all toolkit scores to the default benchmark?");

      if (!shouldReset) {
        return;
      }

      toolkitScoreInputs.forEach((input) => {
        input.value = "3";
      });

      writeToolkitState({});
      updateToolkitScorecard();
    });
  }

  updateToolkitScorecard();
}
