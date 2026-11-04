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
  const doc = DocumentApp.getActiveDocument();
  if (!doc) return null;

  const body = doc.getBody();
  const numChildren = body.getNumChildren();
  const config = actionPointsSettingsManager.load();

  // Load and merge/override configurations
  const peopleConfig = loadMergedPeopleConfig(config.peopleConfig);
  const aliasLookup = buildAliasLookup(peopleConfig);

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

      // --- Dynamic Meeting Date Resolution (Single Pass) ---
      if (!meetingSearchDone && !nextMeetingDate) {
        if (meetingHeaderIndex === null) {
          if (/^(?:Volgende\s+vergadering|Next\s+meetings?)\b/i.test(trimmedText)) {
            meetingHeaderIndex = i;
            meetingHeaderNesting = currentNesting;
          }
        } else {
          if (isListItem && currentNesting > meetingHeaderNesting) {
            const parsed = parseMeetingDate(trimmedText);
            if (parsed) {
              nextMeetingDate = parsed;
              meetingSearchDone = true;
            }
          } else {
            if (isListItem && currentNesting <= meetingHeaderNesting) {
              meetingSearchDone = true;
            } else if (!isListItem && meetingHeaderNesting >= 0) {
              meetingSearchDone = true;
            }
          }
        }
      }
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

  return {
    openTasks,
    completedTasks,
    openMatches,
    completedMatches,
    foundPeople,
    nextMeetingDate,
  };
}
