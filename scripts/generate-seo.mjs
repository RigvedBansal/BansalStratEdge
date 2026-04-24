import { readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

const rootDir = process.cwd();
const siteUrl = "https://bansalstratedge.com";
const siteName = "Bansal StratEdge";
const legalName = "Bansal StratEdge Pvt. Ltd.";
const defaultImagePath = "/assets/blog-finance-banking.png";
const defaultImageAlt = "Bansal StratEdge finance systems and capital strategy visual";
const logoPath = "/assets/bansal-stratedge-logo-wordmark.png";
const orgId = `${siteUrl}/#organization`;
const websiteId = `${siteUrl}/#website`;

const journalFiles = [
  "journal-treasury-mandate.html",
  "journal-capital-confidence.html",
  "journal-board-narratives.html",
  "journal-cfo-dashboards.html",
  "journal-ai-control-tower.html",
  "journal-daily-confidence.html",
  "journal-growth-architecture.html",
];

const toolkitResourceFiles = [
  "cfo-capital-efficiency-checklist.html",
  "capital-scoring-prioritization-matrix.html",
  "cash-conversion-cycle-optimizer.html",
  "cash-flow-runway-planner.html",
  "monthly-variance-bridge.html",
  "variance-narrative-builder.html",
  "driver-based-rolling-forecast.html",
  "driver-decomposition-worksheet.html",
  "headcount-capacity-planning-model.html",
];

const pageOrder = [
  "index.html",
  "finance-systems-toolkit.html",
  "blogs.html",
  "founders.html",
  "workshops.html",
  ...toolkitResourceFiles,
  ...journalFiles,
];

function decodeEntities(value) {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function encodeHtml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function encodeXml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function extract(html, pattern) {
  const match = html.match(pattern);

  if (!match) {
    return "";
  }

  return decodeEntities(match[1]).replace(/\s+/g, " ").trim();
}

function getPagePath(file) {
  if (file === "index.html") {
    return "/";
  }

  if (file === "resources-toolkit.html") {
    return "/finance-systems-toolkit.html";
  }

  return `/${file}`;
}

function getPageUrl(file) {
  return `${siteUrl}${getPagePath(file)}`;
}

function getPageKind(file) {
  if (file === "index.html") {
    return "home";
  }

  if (file === "blogs.html") {
    return "blog-index";
  }

  if (file === "finance-systems-toolkit.html") {
    return "toolkit-index";
  }

  if (file === "founders.html") {
    return "founders";
  }

  if (file === "workshops.html") {
    return "workshops";
  }

  if (file === "resources-toolkit.html") {
    return "redirect";
  }

  if (journalFiles.includes(file)) {
    return "article";
  }

  if (toolkitResourceFiles.includes(file)) {
    return "toolkit-resource";
  }

  return "webpage";
}

function buildBreadcrumbs(file, title) {
  if (file === "index.html") {
    return [{ name: "Home", url: siteUrl }];
  }

  if (file === "blogs.html") {
    return [
      { name: "Home", url: siteUrl },
      { name: "Insights", url: getPageUrl(file) },
    ];
  }

  if (file === "finance-systems-toolkit.html") {
    return [
      { name: "Home", url: siteUrl },
      { name: "Toolkit", url: getPageUrl(file) },
    ];
  }

  if (journalFiles.includes(file)) {
    return [
      { name: "Home", url: siteUrl },
      { name: "Insights", url: `${siteUrl}/blogs.html` },
      { name: title, url: getPageUrl(file) },
    ];
  }

  if (toolkitResourceFiles.includes(file)) {
    return [
      { name: "Home", url: siteUrl },
      { name: "Toolkit", url: `${siteUrl}/finance-systems-toolkit.html` },
      { name: title, url: getPageUrl(file) },
    ];
  }

  return [
    { name: "Home", url: siteUrl },
    { name: title, url: getPageUrl(file) },
  ];
}

function buildOrganizationSchema() {
  return {
    "@type": "Organization",
    "@id": orgId,
    name: siteName,
    legalName,
    url: siteUrl,
    logo: {
      "@type": "ImageObject",
      url: `${siteUrl}${logoPath}`,
    },
    email: "mailto:kamlesh@kamleshbansal.com",
    sameAs: [
      "https://www.kamleshbansal.com/",
      "https://www.linkedin.com/in/kamleshrbansal/",
      "https://www.linkedin.com/in/rigved-kamlesh-bansal/",
    ],
  };
}

function buildWebsiteSchema() {
  return {
    "@type": "WebSite",
    "@id": websiteId,
    url: siteUrl,
    name: siteName,
    publisher: { "@id": orgId },
    inLanguage: "en-IN",
  };
}

function buildPersonSchema(id, name, details = {}) {
  return {
    "@type": "Person",
    "@id": id,
    name,
    ...details,
  };
}

function buildBreadcrumbSchema(file, title) {
  const breadcrumbs = buildBreadcrumbs(file, title);

  return {
    "@type": "BreadcrumbList",
    "@id": `${getPageUrl(file)}#breadcrumbs`,
    itemListElement: breadcrumbs.map((breadcrumb, index) => ({
      "@type": "ListItem",
      position: index + 1,
      name: breadcrumb.name,
      item: breadcrumb.url,
    })),
  };
}

function buildItemList(files, pageMeta) {
  return {
    "@type": "ItemList",
    itemListElement: files.map((file, index) => ({
      "@type": "ListItem",
      position: index + 1,
      item: {
        "@type": "WebPage",
        "@id": `${getPageUrl(file)}#webpage`,
        url: getPageUrl(file),
        name: pageMeta.get(file)?.title ?? file,
      },
    })),
  };
}

function buildPageSchema(file, title, description, pageMeta) {
  const pageUrl = getPageUrl(file);
  const basePage = {
    "@id": `${pageUrl}#webpage`,
    url: pageUrl,
    name: title,
    description,
    inLanguage: "en-IN",
    isPartOf: { "@id": websiteId },
    publisher: { "@id": orgId },
  };

  switch (getPageKind(file)) {
    case "home":
      return {
        "@type": "WebPage",
        ...basePage,
      };

    case "blog-index":
      return {
        "@type": "Blog",
        ...basePage,
        mainEntity: buildItemList(journalFiles, pageMeta),
      };

    case "toolkit-index":
      return {
        "@type": "CollectionPage",
        ...basePage,
        mainEntity: buildItemList(toolkitResourceFiles, pageMeta),
      };

    case "founders": {
      const kamleshId = `${pageUrl}#kamlesh-bansal`;
      const rigvedId = `${pageUrl}#rigved-kamlesh-bansal`;

      return {
        "@type": "AboutPage",
        ...basePage,
        about: [{ "@id": kamleshId }, { "@id": rigvedId }],
      };
    }

    case "workshops":
      return {
        "@type": "ProfilePage",
        ...basePage,
        mainEntity: { "@id": `${pageUrl}#kamlesh-bansal` },
      };

    case "article":
      return {
        "@type": "BlogPosting",
        ...basePage,
        headline: title,
        mainEntityOfPage: pageUrl,
        image: `${siteUrl}${defaultImagePath}`,
        author: { "@id": orgId },
      };

    case "toolkit-resource":
      return {
        "@type": "WebPage",
        ...basePage,
        about: { "@id": `${siteUrl}/finance-systems-toolkit.html#webpage` },
      };

    default:
      return {
        "@type": "WebPage",
        ...basePage,
      };
  }
}

function buildExtraGraph(file) {
  if (file === "founders.html") {
    const pageUrl = getPageUrl(file);

    return [
      buildPersonSchema(`${pageUrl}#kamlesh-bansal`, "Kamlesh Bansal", {
        jobTitle: "Founder",
        sameAs: [
          "https://www.kamleshbansal.com/",
          "https://www.linkedin.com/in/kamleshrbansal/",
        ],
      }),
      buildPersonSchema(`${pageUrl}#rigved-kamlesh-bansal`, "Rigved Kamlesh Bansal", {
        jobTitle: "Co-Founder",
        alumniOf: "BITS Pilani",
        sameAs: ["https://www.linkedin.com/in/rigved-kamlesh-bansal/"],
      }),
    ];
  }

  if (file === "workshops.html") {
    const pageUrl = getPageUrl(file);

    return [
      buildPersonSchema(`${pageUrl}#kamlesh-bansal`, "Kamlesh Bansal", {
        jobTitle: "Enterprise Finance & Capital Strategy Leader",
        sameAs: [
          "https://kamleshbansal.com",
          "https://www.linkedin.com/in/kamleshrbansal",
        ],
      }),
    ];
  }

  return [];
}

function buildStructuredData(file, title, description, pageMeta) {
  if (getPageKind(file) === "redirect") {
    return "";
  }

  const graph = [
    buildOrganizationSchema(),
    buildWebsiteSchema(),
    buildPageSchema(file, title, description, pageMeta),
    buildBreadcrumbSchema(file, title),
    ...buildExtraGraph(file),
  ];

  const json = JSON.stringify(
    {
      "@context": "https://schema.org",
      "@graph": graph,
    },
    null,
    2,
  ).replace(/</g, "\\u003c");

  return `    <script type="application/ld+json">\n${json
    .split("\n")
    .map((line) => `    ${line}`)
    .join("\n")}\n    </script>`;
}

function buildSeoBlock(file, title, description, pageMeta) {
  const pageKind = getPageKind(file);
  const canonicalUrl = getPageUrl(file);
  const imageUrl = `${siteUrl}${defaultImagePath}`;
  const ogType = pageKind === "article" ? "article" : "website";

  if (pageKind === "redirect") {
    return [
      "    <!-- SEO:START -->",
      `    <link rel="canonical" href="${canonicalUrl}" />`,
      '    <meta name="robots" content="noindex,follow" />',
      "    <!-- SEO:END -->",
    ].join("\n");
  }

  const escapedTitle = encodeHtml(title);
  const escapedDescription = encodeHtml(description);
  const structuredData = buildStructuredData(file, title, description, pageMeta);

  return [
    "    <!-- SEO:START -->",
    `    <link rel="canonical" href="${canonicalUrl}" />`,
    '    <meta name="robots" content="index,follow,max-image-preview:large" />',
    '    <meta property="og:locale" content="en_IN" />',
    `    <meta property="og:site_name" content="${encodeHtml(siteName)}" />`,
    `    <meta property="og:type" content="${ogType}" />`,
    `    <meta property="og:title" content="${escapedTitle}" />`,
    `    <meta property="og:description" content="${escapedDescription}" />`,
    `    <meta property="og:url" content="${canonicalUrl}" />`,
    `    <meta property="og:image" content="${imageUrl}" />`,
    `    <meta property="og:image:alt" content="${encodeHtml(defaultImageAlt)}" />`,
    '    <meta name="twitter:card" content="summary_large_image" />',
    `    <meta name="twitter:title" content="${escapedTitle}" />`,
    `    <meta name="twitter:description" content="${escapedDescription}" />`,
    `    <meta name="twitter:image" content="${imageUrl}" />`,
    `    <meta name="twitter:image:alt" content="${encodeHtml(defaultImageAlt)}" />`,
    structuredData,
    "    <!-- SEO:END -->",
  ]
    .filter(Boolean)
    .join("\n");
}

function injectSeoBlock(html, seoBlock) {
  const sanitizedHtml = html
    .replace(/^[ \t]*<link\s+rel="canonical"[^>]*>\s*\n?/gim, "")
    .replace(/^[ \t]*<meta\s+name="robots"[^>]*>\s*\n?/gim, "");

  if (/<!-- SEO:START -->[\s\S]*?<!-- SEO:END -->/m.test(sanitizedHtml)) {
    return sanitizedHtml.replace(/<!-- SEO:START -->[\s\S]*?<!-- SEO:END -->/m, seoBlock);
  }

  const descriptionTag = sanitizedHtml.match(/<meta\s+name="description"[\s\S]*?\/>\n/i);

  if (descriptionTag) {
    return sanitizedHtml.replace(descriptionTag[0], `${descriptionTag[0]}${seoBlock}\n`);
  }

  const viewportTag = sanitizedHtml.match(/<meta\s+name="viewport"[\s\S]*?\/>\n/i);

  if (viewportTag) {
    return sanitizedHtml.replace(viewportTag[0], `${viewportTag[0]}${seoBlock}\n`);
  }

  return sanitizedHtml;
}

function buildSitemap(indexableFiles) {
  const urls = indexableFiles.map((file) => {
    const lastModified = statSync(resolve(rootDir, file)).mtime.toISOString().slice(0, 10);

    return [
      "  <url>",
      `    <loc>${encodeXml(getPageUrl(file))}</loc>`,
      `    <lastmod>${lastModified}</lastmod>`,
      "  </url>",
    ].join("\n");
  });

  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
    ...urls,
    "</urlset>",
    "",
  ].join("\n");
}

const htmlFiles = readdirSync(rootDir)
  .filter((file) => file.endsWith(".html"))
  .sort((left, right) => {
    const leftIndex = pageOrder.indexOf(left);
    const rightIndex = pageOrder.indexOf(right);

    if (leftIndex === -1 && rightIndex === -1) {
      return left.localeCompare(right);
    }

    if (leftIndex === -1) {
      return 1;
    }

    if (rightIndex === -1) {
      return -1;
    }

    return leftIndex - rightIndex;
  });

const pageMeta = new Map();

for (const file of htmlFiles) {
  const html = readFileSync(resolve(rootDir, file), "utf8");
  const title = extract(html, /<title>([\s\S]*?)<\/title>/i);
  const description = extract(html, /<meta\s+name="description"\s+content="([\s\S]*?)"\s*\/?>/i);

  pageMeta.set(file, {
    title,
    description,
  });
}

for (const file of htmlFiles) {
  const filePath = resolve(rootDir, file);
  const html = readFileSync(filePath, "utf8");
  const { title, description } = pageMeta.get(file);
  const seoBlock = buildSeoBlock(file, title, description, pageMeta);
  const nextHtml = injectSeoBlock(html, seoBlock);

  writeFileSync(filePath, nextHtml);
}

const indexableFiles = htmlFiles.filter((file) => getPageKind(file) !== "redirect");

writeFileSync(
  resolve(rootDir, "robots.txt"),
  `User-agent: *\nAllow: /\n\nSitemap: ${siteUrl}/sitemap.xml\n`,
);
writeFileSync(resolve(rootDir, "sitemap.xml"), buildSitemap(indexableFiles));

console.log(`Generated SEO metadata for ${htmlFiles.length} HTML files.`);
