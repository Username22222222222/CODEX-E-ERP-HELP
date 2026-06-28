document.addEventListener('DOMContentLoaded', async () => {
  const toc = document.getElementById('toc');
  const searchInput = document.getElementById('toc-search');
  const pageTitleEl = document.getElementById('doc-page-title');
  const pageTextEl = document.getElementById('doc-page-text');
  let activeLink = null;

  if (!toc) {
    return;
  }

  const setActiveLink = (link) => {
    if (activeLink) {
      activeLink.classList.remove('active');
    }
    activeLink = link;
    if (activeLink) {
      activeLink.classList.add('active');
    }
  };

  const RESX_VIEW_DIRECTORY_FALLBACK = {
    Article: 'Artikel',
    Attachment: 'Anhänge',
    Calendar: 'Kalender',
    Country: 'Länderpakete',
    Crm: 'CRM',
    Device: 'Geräteservice',
    Doc: 'Dok',
    Employee: 'Mitarbeiter',
    FieldService: 'Außendienst',
    Finance: 'Finanzen',
    FormattedSearch: 'Formatierte Suche',
    Helpdesk: 'Helpdesk',
    Intercom: 'Intercom',
    Note: 'Notiz',
    Project: 'Projekt',
    Production: 'Produktion',
    StickyNote: 'Notiz',
    TimeTracking: 'Zeiterfassung',
    Partner: 'Partner',
    Resource: 'Ressource',
    Setup: 'Setup',
    Translation: 'Uebersetzung',
    Warehouse: 'Lager',
    'Formatierte Suche': 'Formatierte Suche',
    'Haft-Notizen': 'Haftnotizen',
    Anhänge: 'Anhänge',
    Projekte: 'Projekt',
    Ressourcen: 'Ressource',
    Länderpakete: 'Länderpakete',
    'Textblöcke': 'Textblock',
    'Extra-Module': 'Extra-Modul',
    Plugins: 'Plugin',
    Haftnotizen: 'Haftnotiz',
    Systemansichten: 'Systemansicht',
    Supporteinstellungen: 'Supporteinstellung',
  };

  const resolveResxViewDirectory = (title) => {
    const resxMap = (window.XERP_RESX_VIEW_DIRECTORIES && typeof window.XERP_RESX_VIEW_DIRECTORIES === 'object')
      ? window.XERP_RESX_VIEW_DIRECTORIES
      : null;

    if (resxMap && typeof resxMap[title] === 'string' && resxMap[title].trim().length > 0) {
      return resxMap[title].trim();
    }

    return RESX_VIEW_DIRECTORY_FALLBACK[title] || title;
  };

  const localizeAnsichtenTitle = (title) => {
    if (!title || typeof title !== 'string') {
      return title;
    }

    const direct = resolveResxViewDirectory(title);
    if (direct !== title) {
      title = direct;
    }

    const tokenReplacements = [
      ['Shipping', 'Versand'],
      ['User', 'Benutzer'],
      ['Procurement', 'Beschaffung'],
      ['Output', 'Ausgabesteuerung'],
      ['FieldService', 'Außendienst'],
    ];

    for (const [from, to] of tokenReplacements) {
      title = title.replace(new RegExp(from, 'g'), to);
    }

    const prefixMap = {
      FormattedSearch: 'Formatierte Suche',
      StickyNote: 'Notiz',
      Note: 'Notiz',
      Device: 'Geräteservice',
      Attachment: 'Anhänge',
      Employee: 'Mitarbeiter',
      FieldService: 'Außendienst',
    };

    for (const [prefix, localized] of Object.entries(prefixMap)) {
      if (title.startsWith(prefix)) {
        return localized + title.slice(prefix.length);
      }
    }

    return title;
  };

  const resolveAnsichtenBranchStem = (title) => {
    const localized = localizeAnsichtenTitle(title || '');
    const stemMap = {
      Anhänge: 'Anhänge',
      Anhang: 'Anhänge',
      Projekte: 'Projekt',
      Projekt: 'Projekt',
      Ressourcen: 'Ressource',
      Ressource: 'Ressource',
      Länderpakete: 'Länderpaket',
      'Textblöcke': 'Textblock',
      'Extra-Module': 'Extra-Modul',
      Plugins: 'Plugin',
      Haftnotizen: 'Haftnotiz',
      Systemansichten: 'Systemansicht',
      Supporteinstellungen: 'Supporteinstellung',
    };

    return stemMap[localized] || localized;
  };

  const displayTitle = (title) => {
    if (!title) {
      return 'X-ERP Dokumentation';
    }

    if (title === 'views') {
      return 'Ansichten';
    }

    if (typeof title === 'string') {
      return localizeAnsichtenTitle(title);
    }

    return title;
  };

  const ARTICLE_EDIT_TAB_LABELS = {
    Overview: 'Übersicht',
    Details: 'Details',
    Sales: 'Verkauf',
    Procurement: 'Beschaffung',
    Purchase: 'Einkauf',
    Text: 'Text',
    Picture: 'Bild',
    Warehousing: 'Lagerung',
    WarehousingHistory: 'Lagerungshistorie',
    Set: 'Set',
    Listing: 'Auflistung',
    Production: 'Produktion',
    Accessories: 'Zubehör',
    Categories: 'Kategorien',
    CatalogNumbers: 'Katalognummern',
    Resource: 'Ressource',
    UsageSet: 'Verwendung bei Sets',
    UsageProduction: 'Verwendung bei Produktion',
    PositionList: 'Positionsliste',
    Revenue: 'Umsatz',
    Contracts: 'Verträge',
    Local: 'Lokal',
    Changelog: 'Änderungsprotokoll',
    Attachments: 'Anlagen',
    ExtraFields: 'Extra-Felder',
  };

  const ARTICLE_EDIT_LABEL_TO_TOKEN = Object.fromEntries(
    Object.entries(ARTICLE_EDIT_TAB_LABELS).map(([token, label]) => [label, token])
  );

  const COMMON_TAB_LABEL_TO_TOKEN = {
    ...ARTICLE_EDIT_LABEL_TO_TOKEN,
    Artikelnummer: 'Number',
    Adresse: 'Address',
    Optionen: 'Options',
    Info: 'Info',
    Historie: 'History',
    Preise: 'Prices',
    Anhaenge: 'Attachments',
    Anhänge: 'Attachments',
    Zusatzfelder: 'ExtraFields',
    Extrafelder: 'ExtraFields',
    Aenderungsprotokoll: 'Changelog',
    Änderungsprotokoll: 'Changelog',
  };

  // Mapping für englische View-Namen zu deutschen
  const VIEW_NAME_LOCALIZE = {
    'ArticleAccessoryEdit': 'Artikel-Zubehör',
    'ArticleCategoryEdit': 'Artikel-Kategorie',
    'ArticleCategoryNameEdit': 'Artikel-Kategorie-Name',
    'ArticleEdit_Procurement': 'Artikel-Beschaffung',
    'ArticleEInvoiceUoMEdit': 'Artikel-E-Rechnung-ME',
    'ArticleGroupEdit': 'Artikel-Gruppe',
    'ArticleHistoryDetail': 'Artikel-Verlauf',
    'ArticleIndividualizationEdit': 'Artikel-Individualisierung',
    'ArticleIndividualizationEdit_DeviceService': 'Artikel-Individualisierung-Service',
    'ArticleIndividualizationQuantityHistoryDetail': 'Artikel-Individualisierung-Menge-Verlauf',
    'ArticleMacroEdit': 'Artikel-Makro',
    'ArticlePriceUnitEdit': 'Artikel-Preiseinheit',
    'ArticleProductEdit': 'Artikel-Produkt',
    'ArticleProductionStepComponentEdit': 'Artikel-Produktionsschritt-Komponente',
    'ArticleProductionStepComponentTypeEdit': 'Artikel-Produktionsschritt-Komponente-Typ',
    'ArticleProductionStepResourceEdit': 'Artikel-Produktionsschritt-Ressource',
    'ArticleQuantityHistoryTableDetail': 'Artikel-Menge-Verlauf-Tabelle',
    'ArticleSalesPriceListEdit': 'Artikel-Verkaufspreis-Liste',
    'ArticleSalesPriceListNameEdit': 'Artikel-Verkaufspreis-Liste-Name',
    'ArticleSalesPriceWizard': 'Artikel-Verkaufspreis-Assistent',
    'ArticleSetEdit': 'Artikel-Set',
    'ArticleTemplateEdit': 'Artikel-Vorlage',
    'ArticleUoMConversionEdit': 'Artikel-ME-Umrechnung',
    'ArticleUoMEdit': 'Artikel-Mengeneinheit',
    'TranslationArticleInfoTextEdit': 'Artikel-Infotext-Übersetzung',
    'TranslationArticleName1Edit': 'Artikel-Name1-Übersetzung',
    'TranslationArticleName2Edit': 'Artikel-Name2-Übersetzung',
    'TranslationArticleOrderTextEdit': 'Artikel-Bestelltext-Übersetzung',
    'TranslationArticleUoMEdit': 'Artikel-ME-Übersetzung',
  };

  const localizeArticleEditTitle = (title, nodePath = '') => {
    // First check if it's a view name that needs translation
    if (VIEW_NAME_LOCALIZE[title]) {
      return VIEW_NAME_LOCALIZE[title];
    }

    if (title === 'Article') {
      return 'Artikel';
    }

    if (!title || !nodePath.includes('Ansichten > Artikel >')) {
      return title;
    }

    const parts = nodePath.split(' > ');

    // Root entry for ArticleEdit
    if (nodePath === 'Ansichten > Artikel > ArticleEdit' && title === 'ArticleEdit') {
      return 'ArticleEdit';
    }

    // Direct children of ArticleEdit (register tabs)
    if (parts.length === 4 && parts[2] === 'ArticleEdit' && !/^ArticleEdit/.test(title)) {
      return `ArticleEdit-${title}`;
    }

    const match = title.match(/^ArticleEdit(?:-(.+))?$/);
    if (!match) {
      return title;
    }

    if (!match[1]) {
      return 'ArticleEdit';
    }

    const token = match[1].replace(/-/g, '');
    const localized = ARTICLE_EDIT_TAB_LABELS[token] || match[1].replace(/-/g, ' ');
    return `ArticleEdit-${localized}`;
  };

  const displayNodeTitle = (title, nodePath = '') => {
    if (nodePath === 'FAQ' || nodePath.startsWith('FAQ >')) {
      return title || 'Verzeichnis';
    }
    if (nodePath === 'Ansichten' || nodePath.startsWith('Ansichten >')) {
      return title || 'Verzeichnis';
    }
    return localizeArticleEditTitle(displayTitle(title), nodePath);
  };

  const promoteAnsichtenTabPages = (nodes) => {
    if (!Array.isArray(nodes) || nodes.length === 0) {
      return;
    }

    const findByPath = (list, targetPath, parentPath = '') => {
      for (const node of list) {
        const currentPath = parentPath ? `${parentPath} > ${node.title}` : node.title;
        if (currentPath === targetPath) {
          return node;
        }
        const childHit = findByPath(node.children || [], targetPath, currentPath);
        if (childHit) {
          return childHit;
        }
      }
      return null;
    };

    const ansichtenNode = findByPath(nodes, 'Ansichten');
    if (!ansichtenNode || !Array.isArray(ansichtenNode.children)) {
      return;
    }

    const classifyTable = (title) => {
      const normalized = (title || '').toLowerCase();

      if (/group/.test(normalized)) return 'Gruppen';
      if (/uom|unit|mengeneinheit|einvoice/.test(normalized)) return 'Mengeneinheiten';
      if (/categor|kategor/.test(normalized)) return 'Kategorien';
      if (/price|preis/.test(normalized)) return 'Preise';
      if (/procurement|purchase|einkauf|beschaffung/.test(normalized)) return 'Beschaffung';
      if (/production|produkt|ressource|component|fertigung/.test(normalized)) return 'Produktion';
      if (/set|accessory|zubehoer|macro|makro/.test(normalized)) return 'Sets-und-Zubehoer';
      if (/individual/.test(normalized)) return 'Individualisierung';
      if (/history|historie|verlauf/.test(normalized)) return 'Historie';
      if (/template|vorlage/.test(normalized)) return 'Vorlagen';

      return 'Weitere-Tabellen';
    };

    ansichtenNode.children.forEach((branch) => {
      if (branch?.title === 'New' || /^New[A-Za-z0-9_\-]*/.test(branch?.title || '')) {
        return;
      }

      const children = Array.isArray(branch?.children) ? branch.children : [];
      if (children.length === 0) {
        return;
      }

      const mainEdit = children.find((child) => /Edit$/.test(child?.title || '') && Array.isArray(child.children) && child.children.length > 0);
      if (!mainEdit) {
        return;
      }

      const promoted = [];
      (mainEdit.children || []).forEach((child) => {
        const tabLabel = (child?.title || '').trim();
        const token = COMMON_TAB_LABEL_TO_TOKEN[tabLabel] || tabLabel.replace(/[^A-Za-z0-9]/g, '');
        if (!token) {
          return;
        }
        promoted.push({ title: `${mainEdit.title}-${token}` });
      });

      if (promoted.length < 3) {
        return;
      }

      const uniquePromotedMap = new Map();
      promoted.forEach((node) => {
        if (!uniquePromotedMap.has(node.title)) {
          uniquePromotedMap.set(node.title, node);
        }
      });

      const groupedTables = new Map();
      children.forEach((child) => {
        const title = (child?.title || '').trim();
        if (!title || title === mainEdit.title || title.startsWith(`${mainEdit.title}-`)) {
          return;
        }

        const folder = classifyTable(title);
        if (!groupedTables.has(folder)) {
          groupedTables.set(folder, []);
        }
        groupedTables.get(folder).push({ title });
      });

      const branchLabel = resolveAnsichtenBranchStem(branch.title || '');
      const folderOrder = [
        'Gruppen',
        'Mengeneinheiten',
        'Kategorien',
        'Preise',
        'Beschaffung',
        'Produktion',
        'Sets-und-Zubehoer',
        'Individualisierung',
        'Historie',
        'Vorlagen',
        'Weitere-Tabellen',
      ];

      const folderNodes = [
        {
          title: `${branchLabel}stammdaten`,
          children: Array.from(uniquePromotedMap.values()),
        },
      ];

      folderOrder.forEach((folderSuffix) => {
        const entries = groupedTables.get(folderSuffix) || [];
        if (entries.length > 0) {
          folderNodes.push({
            title: `${branchLabel}-${folderSuffix}`,
            children: entries,
          });
        }
      });

      branch.children = folderNodes;
    });

    const articleIndex = ansichtenNode.children.findIndex((child) => {
      const title = (child?.title || '').trim();
      return title === 'Artikel' || title === 'Article';
    });
    const articleTranslationIndex = ansichtenNode.children.findIndex((child) => {
      const title = (child?.title || '').trim();
      return title === 'Artikel-Übersetzungen' || title === 'Artikel-Uebersetzungen' || title === 'TranslationArticle';
    });

    if (articleIndex >= 0 && articleTranslationIndex >= 0 && articleIndex !== articleTranslationIndex) {
      const [articleTranslationsNode] = ansichtenNode.children.splice(articleTranslationIndex, 1);
      const articleNode = ansichtenNode.children[articleIndex > articleTranslationIndex ? articleIndex - 1 : articleIndex];
      const articleChildren = Array.isArray(articleNode.children) ? articleNode.children : [];
      const alreadyExists = articleChildren.some((child) => (child?.title || '').trim() === (articleTranslationsNode?.title || '').trim());
      if (!alreadyExists) {
        articleNode.children = [...articleChildren, articleTranslationsNode];
      }
    }

    ansichtenNode.children.sort((a, b) => {
      const aLabel = localizeAnsichtenTitle((a?.title || '').trim());
      const bLabel = localizeAnsichtenTitle((b?.title || '').trim());
      return aLabel.localeCompare(bLabel, 'de', { sensitivity: 'base', numeric: true });
    });
  };

  const isRemovedListPage = (node) => {
    return typeof node?.href === 'string' && /\/de\/help\/views\/[^/]*List\.html$/i.test(node.href);
  };

  const pruneRemovedListPages = (nodes) => {
    if (!Array.isArray(nodes)) {
      return [];
    }

    const isDisallowedView = (title) => {
      return /^(New|Note)(?:$|[A-Z_\-])/.test(title || '') || /^Output(?!Control)/.test(title || '');
    };

    return nodes.reduce((acc, node) => {
      if (!node || isRemovedListPage(node)) {
        return acc;
      }

      if (node?.title === 'New' || /^New[A-Za-z0-9_\-]*/.test(node?.title || '') || isDisallowedView(node?.title || '')) {
        return acc;
      }

      acc.push({
        ...node,
        children: pruneRemovedListPages(node.children),
      });
      return acc;
    }, []);
  };

  const buildFallbackPageText = (node) => {
    const title = displayNodeTitle(node?.title || 'Seite', node?.__nodePath || '');
    const children = Array.isArray(node?.children) ? node.children : [];
    const childLines = children.slice(0, 12).map((c) => `• ${displayNodeTitle(c.title, c.__nodePath || '')}`);

    const intro = `# ${title}\n\nDiese Seite ist in der Verzeichnisstruktur vorhanden, aber es ist aktuell kein eigener Seitentext hinterlegt.`;
    if (childLines.length === 0) {
      return `${intro}\n\nNutzen Sie bei Bedarf die übergeordnete Navigation oder die Suche, um verwandte Inhalte zu öffnen.`;
    }

    return `${intro}\n\nVerfügbare Unterpunkte:\n${childLines.join('\n')}`;
  };

  const escapeHtml = (value) => {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  };

  const renderInlineMarkdown = (value) => {
    return escapeHtml(value)
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      .replace(/`([^`]+)`/g, '<code>$1</code>');
  };

  const renderPageContent = (container, rawText, node) => {
    container.innerHTML = '';

    const sourceTitle = (node?.title || '').trim();
    const localizedTitle = displayNodeTitle(sourceTitle, node?.__nodePath || '');
    const escapedSourceTitle = sourceTitle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const localizedRaw = (rawText || '').trim().replace(
      new RegExp(`^(#{1,6}\\s+)${escapedSourceTitle}(\\s*)$`, 'm'),
      `$1${localizedTitle}$2`
    );
    const raw = (localizedRaw || buildFallbackPageText(node));

    // Normalise: expand markers that were collapsed into single spaces
    let text = raw
      .replace(/(\s|^)(#{1,6})\s/g, '\n$2 ')       // ## Heading
      .replace(/\s---\s/g, '\n---\n')               // horizontal rule
      .replace(/(\s|^)([•*])\s/g, '\n$2 ')          // bullets
      .replace(/(\s|^)(\d+\.\))\s/g, '\n$2 ')       // numbered 1.) 2.)
      .replace(/(\s|^)(\d+\.)\s(?=\S)/g, '\n$2 ')   // numbered 1. 2.
      .replace(/(Taste:\s)/g, '\n$1')               // keyboard shortcut blocks
      .replace(/(Funktion:\s)/g, '\n$1')
      .replace(/\n{3,}/g, '\n\n')
      .trim();

    const lines = text.split('\n').map((l) => l.trim());

    let listEl = null;
    let listType = null;
    let paraLines = [];

    const flushPara = () => {
      if (!paraLines.length) return;
      const val = paraLines.join(' ').trim();
      if (val) {
        const p = document.createElement('p');
        if (currentPageTitle === 'Vorwort') {
          const terms = ['Plugins', 'API-Schnittstellen', 'API', 'Länderpakete'];
          const matches = [];

          terms.forEach((term) => {
            let idx = val.indexOf(term);
            while (idx !== -1) {
              matches.push({ start: idx, end: idx + term.length, term });
              idx = val.indexOf(term, idx + term.length);
            }
          });

          matches.sort((a, b) => a.start - b.start || b.end - a.end);

          const filtered = [];
          let lastEnd = -1;
          matches.forEach((m) => {
            if (m.start >= lastEnd) {
              filtered.push(m);
              lastEnd = m.end;
            }
          });

          if (filtered.length > 0) {
            let cursor = 0;
            filtered.forEach((m) => {
              if (m.start > cursor) {
                p.appendChild(document.createTextNode(val.slice(cursor, m.start)));
              }

              const node = resolveVorwortLink(m.term);
              if (node) {
                const a = document.createElement('a');
                a.href = '#';
                a.className = 'doc-module-link';
                a.textContent = val.slice(m.start, m.end);
                a.addEventListener('click', (ev) => {
                  ev.preventDefault();
                  navigateTo(node);
                });
                p.appendChild(a);
              } else {
                p.appendChild(document.createTextNode(val.slice(m.start, m.end)));
              }

              cursor = m.end;
            });

            if (cursor < val.length) {
              p.appendChild(document.createTextNode(val.slice(cursor)));
            }
          } else {
            p.textContent = val;
          }
        } else {
          p.innerHTML = renderInlineMarkdown(val);
        }
        container.appendChild(p);
      }
      paraLines = [];
    };

    const flushList = () => {
      if (!listEl) return;
      container.appendChild(listEl);
      listEl = null;
      listType = null;
    };

    lines.forEach((line) => {
      if (!line) {
        flushPara();
        return;
      }

      // Headings
      const hm = line.match(/^(#{1,6})\s+(.+)$/);
      if (hm) {
        flushPara(); flushList();
        const lvl = Math.min(6, Math.max(2, hm[1].length + 1));
        const h = document.createElement(`h${lvl}`);
        h.textContent = hm[2].trim();
        container.appendChild(h);
        return;
      }

      // Horizontal rule
      if (line === '---') {
        flushPara(); flushList();
        container.appendChild(document.createElement('hr'));
        return;
      }

      // Markdown image: ![Alt](src)
      const imageMatch = line.match(/^!\[([^\]]*)\]\(([^)]+)\)$/);
      if (imageMatch) {
        flushPara(); flushList();
        const figure = document.createElement('figure');
        const img = document.createElement('img');
        img.src = imageMatch[2].trim();
        img.alt = imageMatch[1].trim() || currentPageTitle || 'Screenshot';
        img.loading = 'lazy';
        const captionText = imageMatch[1].trim();
        if (captionText) {
          const figcaption = document.createElement('figcaption');
          figcaption.textContent = captionText;
          figure.appendChild(img);
          figure.appendChild(figcaption);
        } else {
          figure.appendChild(img);
        }
        container.appendChild(figure);
        return;
      }

      // Keyboard shortcut line: "Taste: …"
      const tasteMatch = line.match(/^(Taste|Funktion):\s*(.+)$/);
      if (tasteMatch) {
        flushPara(); flushList();
        const row = document.createElement('div');
        row.className = 'doc-kbd-row';
        const label = document.createElement('span');
        label.className = 'doc-kbd-label';
        label.textContent = tasteMatch[1] + ':';
        const val = document.createElement('span');
        val.className = tasteMatch[1] === 'Taste' ? 'doc-kbd-key' : 'doc-kbd-desc';
        val.textContent = tasteMatch[2].trim();
        row.appendChild(label);
        row.appendChild(val);
        container.appendChild(row);
        return;
      }

      // Numbered list  1.) or 1.
      const numMatch = line.match(/^(\d+[.)]) +(.+)$/);
      if (numMatch) {
        flushPara();
        if (listType !== 'ol') { flushList(); listEl = document.createElement('ol'); listType = 'ol'; }
        const li = document.createElement('li');
        li.innerHTML = renderInlineMarkdown(numMatch[2].trim());
        listEl.appendChild(li);
        return;
      }

      // Bullet list
      const bulletMatch = line.match(/^[•*-] +(.+)$/);
      if (bulletMatch) {
        flushPara();
        if (listType !== 'ul') { flushList(); listEl = document.createElement('ul'); listType = 'ul'; }
        const li = document.createElement('li');
        const bulletText = bulletMatch[1].trim();
        const moduleNode = (currentPageTitle === 'Vorwort'
          ? resolveVorwortLink(bulletText)
          : null) || resolveAlwaysLinkedView(bulletText);
        if (moduleNode) {
          const a = document.createElement('a');
          a.href = '#';
          a.className = 'doc-module-link';
          a.textContent = bulletText;
          a.addEventListener('click', (ev) => { ev.preventDefault(); navigateTo(moduleNode); });
          li.appendChild(a);
        } else {
          li.innerHTML = renderInlineMarkdown(bulletText);
        }
        listEl.appendChild(li);
        return;
      }

      // Regular text
      flushList();
      paraLines.push(line);
    });

    flushPara();
    flushList();
  };

  // Vorwort-Link-Mapping: Bullet-Label → eindeutiger Baum-Pfad
  const VORWORT_LINK_PATHS = {
    'Verkauf': 'Module > Verkauf',
    'Einkauf': 'Module > Einkauf',
    'Kommissionieren': 'Module > Kommissionieren',
    'Packen': 'Module > Packen',
    'Versand': 'Module > Versand',
    'Produktion': 'Module > Produktion',
    'Beschaffungsassistent': 'Module > Beschaffungsassistent',
    'CRM': 'Module > CRM',
    'Außendienst': 'Module > Außendienst',
    'Geräteservice': 'Module > Geräteservice',
    'Vertragsmanagement': 'Module > Vertragsmanagement',
    'Projektmanagement': 'Module > Projektmanagement',
    'Genehmigungen': 'Module > Genehmigungen',
    'Intercom': 'Module > Intercom',
    'Wiki': 'Module > Wiki',
    'Bulletin': 'Module > Bulletin',
    'Zeiterfassung': 'Module > Zeiterfassung',
    'Archivierung': 'Module > Archivierung',
    'Elektronische Haftnotizen': 'Module > Elektronische Haftnotizen',
    'E-Mail': 'Module > E-Mail',
    'Kalender': 'Module > Kalender',
    'KI-Assistent': 'Module > KI-Assistent',
    'Finanzwesen': 'Module > Finanzwesen',
    'Betriebsdatenerfassung (X-ERP.factory)': 'Portal-Apps > X-ERP.factory',
    'Zeiterfassung (X-ERP.timetracking)': 'Portal-Apps > X-ERP.timetracking',
    'Helpdesk- und Ticketsystem (X-ERP.ticketsystem)': 'Portal-Apps > X-ERP.ticketsystem',
    'Webshop (X-ERP.webshop)': 'Portal-Apps > X-ERP.webshop',
    'Marktplatz (X-ERP.marketplace)': 'Portal-Apps > X-ERP.marketplace',
    'Plugins': 'Branchenlösungen',
    'API': 'API-Schnittstellen',
    'API-Schnittstellen': 'API-Schnittstellen',
    'Länderpakete': 'Länderpakete',
  };

  const ALWAYS_LINK_VIEW_TITLES = new Set([
    'ArticleAccessoryEdit',
    'ArticleCategoryEdit',
    'ArticleCategoryNameEdit',
    'ArticleEdit',
    'ArticleEdit_Procurement',
    'ArticleEInvoiceUoMEdit',
    'ArticleGroupEdit',
    'ArticleHistoryDetail',
    'ArticleIndividualizationEdit',
    'ArticleIndividualizationEdit_DeviceService',
    'ArticleIndividualizationQuantityHistoryDetail',
    'ArticleMacroEdit',
  ]);

  const normalizeVorwortLabel = (value) => {
    return (value || '')
      .toLowerCase()
      .replace(/[‐‑‒–—−]/g, '-')
      .replace(/\s+/g, ' ')
      .replace(/\s*\(\s*/g, ' (')
      .replace(/\s*\)\s*/g, ')')
      .trim();
  };

  const NORMALIZED_VORWORT_LINK_PATHS = Object.fromEntries(
    Object.entries(VORWORT_LINK_PATHS).map(([key, path]) => [normalizeVorwortLabel(key), path])
  );

  let nodeIndex = null;
  let pathIndex = null;

  const buildIndex = (nodes, index, byPath, parentPath = '') => {
    nodes.forEach((n) => {
      const nodePath = parentPath ? `${parentPath} > ${n.title}` : n.title;
      n.__nodePath = nodePath;
      if (!index.has(n.title)) {
        index.set(n.title, n);
      }
      byPath.set(nodePath, n);
      buildIndex(n.children || [], index, byPath, nodePath);
    });
  };

  const reorderAnsichtenViewNodes = (nodes, parentPath = '') => {
    if (!Array.isArray(nodes) || nodes.length === 0) {
      return;
    }

    if (/^Ansichten > .*stammdaten$/i.test(parentPath)) {
      const priorityTokens = ['Overview', 'Details', 'Sales'];

      nodes.sort((a, b) => {
        const aKey = (a?.title || '').trim();
        const bKey = (b?.title || '').trim();

        const aToken = aKey.includes('-') ? aKey.split('-').pop() : aKey;
        const bToken = bKey.includes('-') ? bKey.split('-').pop() : bKey;
        const aPrio = priorityTokens.includes(aToken) ? priorityTokens.indexOf(aToken) : 999;
        const bPrio = priorityTokens.includes(bToken) ? priorityTokens.indexOf(bToken) : 999;

        if (aPrio !== bPrio) {
          return aPrio - bPrio;
        }

        return 0;
      });
    }

    nodes.forEach((node) => {
      const nextPath = parentPath ? `${parentPath} > ${node.title}` : (node.title || '');
      reorderAnsichtenViewNodes(node.children || [], nextPath);
    });
  };

  const resolveVorwortLink = (bulletText) => {
    if (!nodeIndex || !pathIndex) return null;
    const mappedPath =
      VORWORT_LINK_PATHS[bulletText] ||
      NORMALIZED_VORWORT_LINK_PATHS[normalizeVorwortLabel(bulletText)];
    if (mappedPath && pathIndex.has(mappedPath)) {
      return pathIndex.get(mappedPath);
    }

    // Fallback for portal app bullets like "... (X-ERP.factory)"
    const appIdMatch = (bulletText || '').match(/\((X-ERP\.[^)]+)\)/i);
    if (appIdMatch) {
      const appTitle = appIdMatch[1];
      for (const [path, node] of pathIndex.entries()) {
        if (path.endsWith(`> ${appTitle}`) && path.includes('Portal-Apps >')) {
          return node;
        }
      }
    }

    return nodeIndex.get(bulletText) || null;
  };

  const resolveAlwaysLinkedView = (value) => {
    const label = (value || '').trim();
    if (!label || !ALWAYS_LINK_VIEW_TITLES.has(label) || !pathIndex) {
      return null;
    }

    for (const [path, node] of pathIndex.entries()) {
      if (node?.title === label && path.includes('Ansichten > Artikel >')) {
        return node;
      }
    }

    return nodeIndex ? nodeIndex.get(label) || null : null;
  };

  let currentPageTitle = '';

  const navigateTo = (node) => {
    // find the DOM <a> whose text matches node.title and click it
    const allLinks = toc.querySelectorAll('a');
    for (const a of allLinks) {
      if (a.dataset.nodePath === node.__nodePath) {
        // expand all ancestor <li> nodes
        let parent = a.parentElement;
        while (parent && parent !== toc) {
          if (parent.tagName === 'LI') parent.classList.add('expanded');
          parent = parent.parentElement;
        }
        showPage(node, a);
        a.scrollIntoView({ block: 'nearest' });
        return;
      }
    }
    // fallback: just show content
    showPage(node, null);
  };

  const expandLinkParents = (link) => {
    let parent = link?.parentElement;
    while (parent && parent !== toc) {
      if (parent.tagName === 'LI') parent.classList.add('expanded');
      parent = parent.parentElement;
    }
  };

  const normalizePath = (value) => {
    const path = (value || '').replace(/^https?:\/\/[^/]+/i, '').split(/[?#]/)[0].toLowerCase();
    return path.endsWith('/') ? path : `${path}/`;
  };

  const isHelpShellRootPath = (value) => {
    const path = normalizePath(value);
    return path === '/de/help/' || path === '/de/help/index.html/';
  };

  const activateCurrentLocation = () => {
    const currentPath = normalizePath(window.location.pathname);
    for (const link of toc.querySelectorAll('a')) {
      const node = pathIndex?.get(link.dataset.nodePath || '');
      if (node?.path && normalizePath(node.path) === currentPath) {
        expandLinkParents(link);
        setActiveLink(link);
        return true;
      }
    }

    return false;
  };

  const showPage = (node, link, options = {}) => {
    if (!options.inline && typeof node?.path === 'string' && node.path.startsWith('/de/help/')) {
      window.location.href = node.path;
      return;
    }

    const title = (node?.title || '').trim();
    const path = node?.__nodePath || '';
    const isViewsBranch = path.startsWith('Ansichten >');
    const isNestedViewsEntry = /^Ansichten\s*>\s*[^>]+\s*>/.test(path);
    const isTechnicalViewId = /^[A-Za-z][A-Za-z0-9_]*$/.test(title)
      && (/[a-z][A-Z]/.test(title) || /_/.test(title) || /[A-Z]{2,}/.test(title));
    const isListViewId = /List(?:$|[_-])/.test(title);
    const isLikelyViewPage = isTechnicalViewId && !isListViewId;

    if (isViewsBranch && isNestedViewsEntry && isLikelyViewPage) {
      window.location.href = `./views/${title}.html`;
      return;
    }

    currentPageTitle = node.title || '';
    if (pageTitleEl) {
      pageTitleEl.textContent = displayNodeTitle(node.title || 'X-ERP Dokumentation', node.__nodePath || '');
    }
    if (pageTextEl) {
      renderPageContent(pageTextEl, node.pageTitle || '', node);
    }
    setActiveLink(link);
  };

  const buildTree = (nodes, parentUl, level = 0) => {
    nodes.forEach((node) => {
      const title = (node?.title || '').trim();
      const nodePath = node?.__nodePath || '';
      const isTechnicalViewId = /^[A-Za-z][A-Za-z0-9_]*$/.test(title)
        && (/[a-z][A-Z]/.test(title) || /_/.test(title) || /[A-Z]{2,}/.test(title));
      const isListViewId = /List(?:$|[_-])/.test(title);
      const isAnsichtenNestedEntry = /^Ansichten\s*>\s*[^>]+\s*>/i.test(nodePath);
      const isAnsichtenViewPage = isAnsichtenNestedEntry && isTechnicalViewId && !isListViewId;
      const isExcelAnsichtenNode = typeof node.sourceRow === 'number';
      const hasChildren = !(isAnsichtenViewPage && !isExcelAnsichtenNode) && Array.isArray(node.children) && node.children.length > 0;
      const li = document.createElement('li');
      li.className = hasChildren ? 'toc-dir-item' : 'toc-leaf-page';

      const link = document.createElement('a');
      const nodeHref = typeof node.path === 'string' && node.path.startsWith('/de/help/')
        ? node.path
        : '#';
      link.href = nodeHref;
      link.className = hasChildren ? 'toc-dir-link' : 'toc-page-link';
      link.textContent = displayNodeTitle(node.title || 'Verzeichnis', node.__nodePath || '');
      link.dataset.nodePath = node.__nodePath || '';
      link.addEventListener('click', (event) => {
        if (nodeHref !== '#') {
          event.preventDefault();
          setActiveLink(link);
          if (window.location.pathname !== nodeHref) {
            window.location.href = nodeHref;
          }
          return;
        }
        event.preventDefault();
        showPage(node, link);
      });

      li.appendChild(link);

      if (hasChildren) {
        const childUl = document.createElement('ul');
        buildTree(node.children, childUl, level + 1);
        li.appendChild(childUl);
      }

      if (!hasChildren && level === 0 && (node.title === 'Vorwort' || node.title === 'Inhaltsverzeichnis')) {
        li.classList.add('toc-top-page');
      }

      parentUl.appendChild(li);
    });
  };

  const attachToggles = (root) => {
    root.querySelectorAll('li').forEach((li) => {
      const childList = li.querySelector(':scope > ul');
      if (!childList) {
        return;
      }

      const toggle = document.createElement('button');
      toggle.type = 'button';
      toggle.className = 'toc-toggle';
      toggle.setAttribute('aria-label', 'Unterverzeichnis ein-/ausklappen');
      li.insertBefore(toggle, li.firstChild);

      toggle.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
        li.classList.toggle('expanded');
      });
    });
  };

  const resetTree = () => {
    toc.querySelectorAll('li').forEach((li) => {
      li.style.display = '';
      li.classList.remove('expanded');
    });
  };

  const filterNode = (li, term) => {
    const ownLink = li.querySelector(':scope > a');
    const ownText = ownLink ? ownLink.textContent.toLowerCase() : '';
    const ownMatch = ownText.includes(term);

    let childMatch = false;
    li.querySelectorAll(':scope > ul > li').forEach((child) => {
      if (filterNode(child, term)) {
        childMatch = true;
      }
    });

    const visible = ownMatch || childMatch;
    li.style.display = visible ? '' : 'none';

    if (term.length > 0) {
      li.classList.toggle('expanded', childMatch || ownMatch);
    }

    return visible;
  };

  const loadTreeData = async () => {
    if (Array.isArray(window.XERP_XLSX_TREE)) {
      return window.XERP_XLSX_TREE;
    }

    const candidates = [
      '/de/help/xlsx-tree.json',
      './xlsx-tree.json',
      'xlsx-tree.json'
    ];

    let lastError = null;
    for (const url of candidates) {
      try {
        const response = await fetch(url, { cache: 'no-store' });
        if (!response.ok) {
          throw new Error(`HTTP ${response.status} fuer ${url}`);
        }
        const data = await response.json();
        if (!Array.isArray(data)) {
          throw new Error(`Ungueltiges Format in ${url}`);
        }
        return data;
      } catch (err) {
        lastError = err;
      }
    }

    throw lastError || new Error('xlsx-tree.json konnte nicht geladen werden');
  };

  const cloneNode = (node) => ({
    ...node,
    children: Array.isArray(node.children) ? node.children.map(cloneNode) : undefined,
  });

  const loadAnsichtenBranch = async () => {
    if (window.XERP_ANSICHTEN_TREE && typeof window.XERP_ANSICHTEN_TREE === 'object') {
      return window.XERP_ANSICHTEN_TREE;
    }

    const candidates = [
      '/de/help/ansichten-tree.json',
      './ansichten-tree.json',
      'ansichten-tree.json'
    ];

    for (const url of candidates) {
      try {
        const response = await fetch(url, { cache: 'no-store' });
        if (!response.ok) {
          continue;
        }
        const data = await response.json();
        if (data && typeof data === 'object' && data.title === 'Ansichten') {
          return data;
        }
      } catch (err) {
        // Fallback candidates are tried below.
      }
    }

    return null;
  };

  const replaceAnsichtenBranchFromExcel = (nodes, ansichtenTree) => {
    if (!ansichtenTree || typeof ansichtenTree !== 'object') {
      return nodes;
    }

    const replacement = cloneNode(ansichtenTree);
    const clonedNodes = Array.isArray(nodes) ? nodes.map(cloneNode) : [];
    const index = clonedNodes.findIndex((node) => (node?.title || '').trim() === 'Ansichten');
    if (index >= 0) {
      clonedNodes[index] = replacement;
      return clonedNodes;
    }

    clonedNodes.push(replacement);
    return clonedNodes;
  };

  try {
    const treeData = await loadTreeData();
    const ansichtenTree = await loadAnsichtenBranch();
    let nodes = pruneRemovedListPages(Array.isArray(treeData) ? treeData : []);
    nodes = replaceAnsichtenBranchFromExcel(nodes, ansichtenTree);
    if (!ansichtenTree) {
      promoteAnsichtenTabPages(nodes);
      reorderAnsichtenViewNodes(nodes);
    }
    nodeIndex = new Map();
    pathIndex = new Map();
    buildIndex(nodes, nodeIndex, pathIndex);
    buildTree(nodes, toc);
    attachToggles(toc);

    const firstNode = nodes[0];
    const firstLink = toc.querySelector(':scope > li > a');
    const isRootDirectoryPage = isHelpShellRootPath(window.location.pathname);
    if (!activateCurrentLocation() && firstNode && firstLink && !isRootDirectoryPage) {
      showPage(firstNode, firstLink);
    }
  } catch (error) {
    console.error('Fehler beim Laden der Verzeichnisstruktur:', error);
    const li = document.createElement('li');
    li.textContent = 'Fehler beim Laden der Verzeichnisstruktur.';
    toc.appendChild(li);
  }

  if (!searchInput) {
    return;
  }

  searchInput.addEventListener('input', () => {
    const term = searchInput.value.trim().toLowerCase();
    if (!term) {
      resetTree();
      return;
    }

    toc.querySelectorAll(':scope > li').forEach((li) => {
      filterNode(li, term);
    });
  });
});
