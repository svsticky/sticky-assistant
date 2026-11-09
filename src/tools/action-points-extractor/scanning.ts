import { PeopleConfigEntry, validatePeopleConfig } from './config';
import { actionPointsSettingsManager } from './settings';

export interface ActionPointOccurrence {
  originalName: string;
  assignees: string[];
  action: string;
  date: string | null;
  isBold: boolean;
  location: {
    childIndex: number;
    matchIndex: number;
    matchLength: number;
  };
  originalText: string;
}

export interface ActionPointTask {
  person: string;
  action: string;
  date: string | null;
  order: number | null;
  childIndex?: number;
  originalText?: string;
}

export interface ActionPointsScanResult {
  openTasks: ActionPointTask[];
  completedTasks: ActionPointTask[];
  openMatches: ActionPointOccurrence[];
  completedMatches: ActionPointOccurrence[];
  foundPeople: Record<string, boolean>;
  nextMeetingDate: string | null;
}

function buildAliasLookup(peopleConfig: Record<string, PeopleConfigEntry>) {
  const lookup: Record<string, string> = {};
  Object.keys(peopleConfig || {}).forEach((personName) => {
    const entry = peopleConfig[personName] || {};
    lookup[personName.toLowerCase()] = personName;
    const aliases = Array.isArray(entry.aliases) ? entry.aliases : [];
    aliases.forEach((alias: string) => {
      if (alias && typeof alias === 'string') {
        lookup[alias.toLowerCase()] = personName;
      }
    });
  });
  return lookup;
}

/**
 * Splits assignees by "and" or "&", maps them to correct alias keys, and
 * supports English "everyone" and Dutch "iedereen" assignee triggers.
 */
function resolveAssignees(
  namePart: string,
  peopleConfig: Record<string, PeopleConfigEntry>,
  aliasLookup: Record<string, string>
): string[] {
  const parts = namePart
    .split(/\s+and\s+|\s*&\s*|\s+en\s+|\s*,\s+/gi)
    .map((p) => p.trim())
    .filter(Boolean);
  const resolved: string[] = [];

  parts.forEach((part) => {
    const key = part.toLowerCase();
    if (key === 'everyone' || key === 'iedereen') {
      const people = Object.keys(peopleConfig || {});
      if (people.length > 0) {
        resolved.push(...people);
      } else {
        resolved.push(part);
      }
    } else {
      const person = aliasLookup[key];
      resolved.push(person ? person : part);
    }
  });

  return [...new Set(resolved)];
}

function getOrderForPerson(personName: string, peopleConfig: Record<string, PeopleConfigEntry>) {
  const entry = peopleConfig[personName] || {};
  return typeof entry.order === 'number' ? entry.order : null;
}

/**
 * Shared comparator for ordering action points consistently.
 * Sorts by ascending order property (with unconfigured tasks at the end),
 * tiebreaking alphabetically by person name.
 */
export function compareActionPoints(
  a: { order?: number | null; person: string },
  b: { order?: number | null; person: string }
): number {
  const aOrder = a.order !== null && a.order !== undefined ? a.order : Number.MAX_SAFE_INTEGER;
  const bOrder = b.order !== null && b.order !== undefined ? b.order : Number.MAX_SAFE_INTEGER;
  if (aOrder !== bOrder) return aOrder - bOrder;
  return a.person.localeCompare(b.person);
}

export const MEETING_HEADER_REGEX =
  /^\s*(?:(?:\d+|[a-zA-Z])[.)]\s*|[•\-*]\s*|#{1,6}\s*)*(?:Volgende\s+(?:vergadering|meeting|bijeenkomst)|Next\s+meetings?)\b/i;

export function parseMeetingDate(text: string): string | null {
  if (!text) return null;

  const isoMatch = text.match(/\b(\d{4})[-/](\d{2})[-/](\d{2})\b/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }

  const dmyMatch = text.match(/\b(\d{2})[-/](\d{2})[-/](\d{4})\b/);
  if (dmyMatch) {
    return `${dmyMatch[3]}-${dmyMatch[2]}-${dmyMatch[1]}`;
  }

  return null;
}

/**
 * Searches the folder directory tree of the active document to find a "sticky-assistant.json"
 * containing folder-level overrides of the team people configuration.
 */
function loadMergedPeopleConfig(
  globalConfig: Record<string, PeopleConfigEntry>
): Record<string, PeopleConfigEntry> {
  try {
    const doc = DocumentApp.getActiveDocument();
    if (!doc) return globalConfig;

    const docId = doc.getId();
    const file = DriveApp.getFileById(docId);
    const parents = file.getParents();
    if (parents.hasNext()) {
      const parent = parents.next();
      const configFiles = parent.getFilesByName('sticky-assistant.json');
      if (configFiles.hasNext()) {
        const configFile = configFiles.next();
        const content = configFile.getBlob().getDataAsString();
        const parsed = JSON.parse(content);

        // Accept a nested configuration under toolId or general peopleConfig
        const toolOverride = parsed.actionPointsExtractor || {};
        const localPeople =
          toolOverride.peopleConfig || parsed.peopleConfig || parsed.people || parsed;
        const validation = validatePeopleConfig(localPeople);
        if (validation.ok) {
          console.log(
            'Successfully loaded sticky-assistant.json override from parent Drive folder'
          );
          return localPeople as Record<string, PeopleConfigEntry>;
        } else {
          console.warn('sticky-assistant.json found but invalid: ' + validation.message);
        }
      }
    }
  } catch (err) {
    console.warn('Failed to load local sticky-assistant.json override: ', err);
  }
  return globalConfig;
}

export function scanActionPointsDocument(): ActionPointsScanResult | null {
  const startTime = Date.now();
  const doc = DocumentApp.getActiveDocument();
  if (!doc) {
    console.warn('[scanActionPointsDocument] No active document found.');
    return null;
  }

  const docId = typeof doc.getId === 'function' ? doc.getId() : 'unknown';
  const docName = typeof doc.getName === 'function' ? doc.getName() : 'unknown';
  const body = doc.getBody();
  const numChildren = body.getNumChildren();
  console.log(
    `[scanActionPointsDocument] Scanning document "${docName}" (ID: ${docId}, ${numChildren} top-level elements)...`
  );
  const config = actionPointsSettingsManager.load();

  // Load and merge/override configurations
  const peopleConfig = loadMergedPeopleConfig(config.peopleConfig);
  const aliasLookup = buildAliasLookup(peopleConfig);
  console.log(
    `[scanActionPointsDocument] Loaded people config with ${Object.keys(peopleConfig).length} person entries.`
  );

  // Expect standard YYYY-MM-DD or DD-MM-YYYY date format standard: [2026-05-27] or [27-05-2026]
  const dateRegex = /\[(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})\]/;

  const openTasks: ActionPointTask[] = [];
  const completedTasks: ActionPointTask[] = [];
  const openMatches: ActionPointOccurrence[] = [];
  const completedMatches: ActionPointOccurrence[] = [];
  const foundPeople: Record<string, boolean> = {};

  let meetingHeaderIndex: number | null = null;
  let meetingHeaderNesting = -1;
  let nextMeetingDate: string | null = null;
  let meetingSearchDone = false;

  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    const type = child.getType();

    if (type !== DocumentApp.ElementType.PARAGRAPH && type !== DocumentApp.ElementType.LIST_ITEM) {
      continue;
    }

    const text = (child as any).getText();
    const trimmedText = text.trim();

    if (trimmedText.length > 0) {
      const isListItem = type === DocumentApp.ElementType.LIST_ITEM;
      const currentNesting = isListItem ? (child as any).getNestingLevel() : -1;

      // --- Dynamic Meeting Date Resolution ---
      if (!meetingSearchDone && !nextMeetingDate) {
        if (meetingHeaderIndex !== null) {
          // Check if this child under the meeting header candidate contains a valid date
          const parsed = parseMeetingDate(trimmedText);
          if (parsed) {
            nextMeetingDate = parsed;
            meetingSearchDone = true;
            console.log(
              `[scanActionPointsDocument] Extracted next meeting date "${nextMeetingDate}" from child index ${i}: "${trimmedText}".`
            );
          } else {
            // Check if this candidate section has ended without finding a date
            const isNextNumberedSection = /^\s*(?:\d+[.)]|#{1,6})\s*[A-Za-z]/i.test(trimmedText);
            const isHigherOrEqualListLevel =
              meetingHeaderNesting >= 0 && isListItem && currentNesting <= meetingHeaderNesting;
            const exceededFirstChildren = i - meetingHeaderIndex >= 2;

            if (isNextNumberedSection || isHigherOrEqualListLevel || exceededFirstChildren) {
              console.log(
                `[scanActionPointsDocument] Candidate at child index ${meetingHeaderIndex} did not contain a valid meeting date (ended at child index ${i}: "${trimmedText}"). Continuing search for next meeting section...`
              );
              // Reset candidate header so we continue scanning the rest of the document
              meetingHeaderIndex = null;
              meetingHeaderNesting = -1;
            }
          }
        }

        // Check if current child is a new candidate header
        if (meetingHeaderIndex === null && MEETING_HEADER_REGEX.test(trimmedText)) {
          meetingHeaderIndex = i;
          meetingHeaderNesting = currentNesting;
          console.log(
            `[scanActionPointsDocument] Meeting section candidate found at child index ${i}: "${trimmedText}"`
          );

          // Check if meeting date is on the same line as the header
          const sameLineDate = parseMeetingDate(trimmedText);
          if (sameLineDate) {
            nextMeetingDate = sameLineDate;
            meetingSearchDone = true;
            console.log(
              `[scanActionPointsDocument] Extracted next meeting date "${nextMeetingDate}" from header line at child index ${i}.`
            );
          }
        }
      }
    }

    if (!/\bAP\s+/i.test(text)) {
      continue;
    }

    const regex = /\bAP\s+([^:]+):\s+(.+)/gi;
    let match: RegExpExecArray | null;

    while ((match = regex.exec(text)) !== null) {
      const textElement = (child as any).editAsText();
      const isBold = textElement.isBold(match.index) && textElement.isBold(match.index + 1);

      let namePart = match[1].trim();
      let actionPart = match[2].trim();
      let taskDate: string | null = null;

      let dateMatch = namePart.match(dateRegex);
      if (dateMatch) {
        taskDate = parseMeetingDate(dateMatch[1]);
        namePart = namePart.replace(dateRegex, '').trim();
      } else {
        dateMatch = actionPart.match(dateRegex);
        if (dateMatch) {
          taskDate = parseMeetingDate(dateMatch[1]);
          actionPart = actionPart.replace(dateRegex, '').trim();
        }
      }

      if (actionPart.length > 0) {
        actionPart = actionPart.charAt(0).toUpperCase() + actionPart.slice(1);
      }

      const assignees = resolveAssignees(namePart, peopleConfig, aliasLookup);
      console.log(
        `[scanActionPointsDocument] Child ${i}: Matched AP [${assignees.join(', ')}] "${actionPart}" (due: ${taskDate || 'none'}, status: ${isBold ? 'OPEN' : 'COMPLETED'})`
      );

      const matchObj: ActionPointOccurrence = {
        originalName: namePart,
        assignees,
        action: actionPart,
        date: taskDate,
        isBold,
        location: {
          childIndex: i,
          matchIndex: match.index,
          matchLength: match[0].length,
        },
        originalText: match[0],
      };

      assignees.forEach((assignee) => {
        foundPeople[assignee] = true;
        const isConfigured = Boolean(peopleConfig[assignee] || aliasLookup[assignee.toLowerCase()]);
        if (
          !isConfigured &&
          assignee.toLowerCase() !== 'everyone' &&
          assignee.toLowerCase() !== 'iedereen'
        ) {
          console.log(
            `[scanActionPointsDocument] Notice: Assignee "${assignee}" is not configured in settings.`
          );
        }

        const task: ActionPointTask = {
          person: assignee,
          action: actionPart,
          date: taskDate,
          order: getOrderForPerson(assignee, peopleConfig),
          childIndex: i,
          originalText: match![0],
        };
        if (isBold) {
          openTasks.push(task);
        } else {
          completedTasks.push(task);
        }
      });

      if (isBold) {
        openMatches.push(matchObj);
      } else {
        completedMatches.push(matchObj);
      }
    }
  }

  const durationMs = Date.now() - startTime;
  console.log(
    `[scanActionPointsDocument] Scan completed in ${durationMs}ms. Found: ${openTasks.length} open tasks (${openMatches.length} occurrences), ${completedTasks.length} completed tasks (${completedMatches.length} occurrences), ${Object.keys(foundPeople).length} detected people: [${Object.keys(foundPeople).join(', ')}], nextMeetingDate: ${nextMeetingDate || 'none'}`
  );

  return {
    openTasks,
    completedTasks,
    openMatches,
    completedMatches,
    foundPeople,
    nextMeetingDate,
  };
}
