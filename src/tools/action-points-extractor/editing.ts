import { actionPointsSettingsManager } from './settings';
import { PeopleConfigEntry } from './config';
import { ActionPointTask, compareActionPoints, scanActionPointsDocument } from './scanning';

export interface ActionPointsOperationResult {
  ok: boolean;
  message: string;
  successCount?: number;
  addedCount?: number;
}

function formatActionPoint(nameText: string, actionText: string, dateText: string | null) {
  const dateSuffix = dateText ? ` [${dateText}]` : '';
  return `AP ${nameText}: ${actionText}${dateSuffix}`;
}

export function formatTodoistExportTask(
  task: ActionPointTask,
  peopleConfig: Record<string, PeopleConfigEntry>,
  defaultDueDate?: string | null
): string {
  const parts: string[] = ['[AP]', task.action];

  const entry = peopleConfig[task.person];
  if (entry?.todoist_section) {
    parts.push(`/${entry.todoist_section}`);
  }

  const date = task.date || defaultDueDate;
  if (date) {
    parts.push(date);
  }

  const todoistId = entry?.todoist_id || task.person;
  if (todoistId) {
    parts.push(`+${todoistId}`);
  }

  return parts.join(' ');
}

export function getNewDetectedPeople(
  foundPeople: Record<string, boolean>,
  peopleConfig: Record<string, PeopleConfigEntry>
): string[] {
  const existingNames = new Set<string>();
  Object.keys(peopleConfig || {}).forEach((name) => {
    existingNames.add(name.toLowerCase().trim());
    const entry = peopleConfig[name];
    if (entry && Array.isArray(entry.aliases)) {
      entry.aliases.forEach((alias) => {
        if (alias && typeof alias === 'string') {
          existingNames.add(alias.toLowerCase().trim());
        }
      });
    }
  });

  return Object.keys(foundPeople || {}).filter((name) => {
    const trimmed = name.trim();
    if (!trimmed || trimmed.toLowerCase() === 'everyone' || trimmed.toLowerCase() === 'iedereen') {
      return false;
    }
    return !existingNames.has(trimmed.toLowerCase());
  });
}

export function sendToTodoistLogic(e: any): ActionPointsOperationResult {
  console.log('[sendToTodoistLogic] Initiating Todoist sync...');
  let tasks: any[] = [];
  let defaultDueDate: string | null = null;
  if (Array.isArray(e)) {
    tasks = e;
  } else if (e && e.parameters) {
    if (e.parameters.tasksJson) {
      try {
        tasks = JSON.parse(e.parameters.tasksJson);
      } catch (err) {
        console.warn('[sendToTodoistLogic] Failed to parse tasksJson parameter: ', err);
      }
    }
    if (e.parameters.nextMeetingDate) {
      defaultDueDate = e.parameters.nextMeetingDate;
    }
  }

  if (!Array.isArray(tasks) || tasks.length === 0) {
    console.log('[sendToTodoistLogic] No tasks in parameters; scanning active document...');
    const scanResult = scanActionPointsDocument();
    if (scanResult) {
      tasks = scanResult.openTasks;
      defaultDueDate = scanResult.nextMeetingDate;
      console.log(
        `[sendToTodoistLogic] Fresh document scan found ${tasks.length} open tasks, defaultDueDate: ${defaultDueDate || 'none'}.`
      );
    }
  }

  const config = actionPointsSettingsManager.load(e);
  if (!config.todoistToken || !config.todoistProjectId) {
    console.warn('[sendToTodoistLogic] Missing Todoist Token or Project ID in settings.');
    return { ok: false, message: 'Please add your Todoist Token and Project ID in Settings.' };
  }

  if (tasks.length === 0) {
    console.log('[sendToTodoistLogic] No open action points to sync.');
    return { ok: true, message: 'No open action points to sync.', successCount: 0 };
  }

  console.log(
    `[sendToTodoistLogic] Syncing ${tasks.length} tasks to Todoist project ${config.todoistProjectId}...`
  );
  let successCount = 0;
  const url = 'https://api.todoist.com/rest/v2/tasks';

  tasks.forEach((task: any) => {
    const content = `${task.person}: ${task.action}`;
    const payloadObj: any = { content, project_id: config.todoistProjectId };

    const effectiveDate = task.date || defaultDueDate;
    if (effectiveDate) {
      payloadObj.due_date = effectiveDate;
    }

    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        headers: {
          Authorization: 'Bearer ' + config.todoistToken,
          'Content-Type': 'application/json',
        },
        muteHttpExceptions: true,
        payload: JSON.stringify(payloadObj),
      });

      if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
        successCount++;
      } else {
        console.warn(
          `[sendToTodoistLogic] Task failed: HTTP ${response.getResponseCode()} - ${response.getContentText()}`
        );
      }
    } catch (err) {
      console.error('[sendToTodoistLogic] Network error syncing task: ', err);
    }
  });

  const message = `Synced ${successCount} of ${tasks.length} tasks to Todoist.`;
  console.log(`[sendToTodoistLogic] Completed: ${message}`);
  return { ok: true, message, successCount };
}

export function applyDocumentActionsLogic(e: any): ActionPointsOperationResult {
  console.log('[applyDocumentActionsLogic] Processing document action updates...');
  let parsed: any = {};
  if (Array.isArray(e)) {
    parsed = e;
  } else if (e && e.parameters && (e.parameters.matchesJson || e.parameters.resultJson)) {
    try {
      parsed = JSON.parse(e.parameters.matchesJson || e.parameters.resultJson);
    } catch (err) {
      console.warn('[applyDocumentActionsLogic] Failed to parse matchesJson parameter: ', err);
    }
  }

  const formInput = (e && e.formInput) || {};
  const formInputs = (e && e.formInputs) || {};

  const addToTop =
    (e && e.parameters && e.parameters.addToTop === 'true') ||
    formInput.addToTopAction === 'addToTop' ||
    (formInputs.addToTopAction && formInputs.addToTopAction.indexOf('addToTop') !== -1);

  const replaceInPlace =
    (e && e.parameters && e.parameters.replaceInPlace === 'true') ||
    formInput.replaceInPlaceAction === 'replaceInPlace' ||
    (formInputs.replaceInPlaceAction &&
      formInputs.replaceInPlaceAction.indexOf('replaceInPlace') !== -1);

  console.log(
    `[applyDocumentActionsLogic] Options selected: addToTop=${Boolean(addToTop)}, replaceInPlace=${Boolean(replaceInPlace)}`
  );

  const doc = DocumentApp.getActiveDocument();
  if (!doc) {
    console.warn('[applyDocumentActionsLogic] No active Google Doc found.');
    return { ok: false, message: 'No active Google Doc found.' };
  }

  if (!addToTop && !replaceInPlace) {
    console.log('[applyDocumentActionsLogic] No checkbox selected.');
    return { ok: true, message: 'No actions selected. Please check at least one box.' };
  }

  const body = doc.getBody();

  let openMatches = parsed.openMatches;
  let completedMatches = parsed.completedMatches;

  if (!openMatches && !completedMatches) {
    console.log(
      '[applyDocumentActionsLogic] No matches passed in parameters; scanning active document...'
    );
    const scanResult = scanActionPointsDocument();
    if (scanResult) {
      openMatches = scanResult.openMatches;
      completedMatches = scanResult.completedMatches;
      console.log(
        `[applyDocumentActionsLogic] Fresh scan found ${openMatches.length} open and ${completedMatches.length} completed matches.`
      );
    } else {
      openMatches = [];
      completedMatches = [];
    }
  } else {
    openMatches = openMatches || [];
    completedMatches = completedMatches || [];
    console.log(
      `[applyDocumentActionsLogic] Using passed matches: ${openMatches.length} open, ${completedMatches.length} completed.`
    );
  }

  // Helper function to replace matches
  const doReplace = (matches: any[], strikeThrough: boolean) => {
    matches.forEach((m: any) => {
      try {
        if (m.location && typeof m.location.childIndex === 'number') {
          const child = body.getChild(m.location.childIndex);
          if (
            child &&
            (child.getType() === DocumentApp.ElementType.PARAGRAPH ||
              child.getType() === DocumentApp.ElementType.LIST_ITEM)
          ) {
            const textElement = (child as any).editAsText();
            const start = m.location.matchIndex;
            const end = m.location.matchIndex + m.location.matchLength - 1;
            textElement.setStrikethrough(start, end, strikeThrough);
            return;
          }
        }
      } catch (err) {
        console.warn('Failed to format by location, falling back to text search: ', err);
      }

      const matchText = m.originalText || m.matchText;
      if (!matchText) return;
      const searchResult = body.findText(matchText);
      if (searchResult) {
        const textElement = searchResult.getElement().asText();
        const start = searchResult.getStartOffset();
        const end = searchResult.getEndOffsetInclusive();
        textElement.setStrikethrough(start, end, strikeThrough);
      }
    });
  };

  // Perform strike-through modifications
  let replaceCount = 0;
  if (replaceInPlace && completedMatches && completedMatches.length > 0) {
    console.log(
      `[applyDocumentActionsLogic] Formatting ${completedMatches.length} completed matches with strikethrough in-place...`
    );
    try {
      doReplace(completedMatches, true);
      replaceCount = completedMatches.length;
      console.log('[applyDocumentActionsLogic] Strikethrough completed.');
    } catch (err) {
      console.error('[applyDocumentActionsLogic] Strikethrough error: ', err);
    }
  }

  let topCount = 0;
  if (addToTop) {
    console.log(
      `[applyDocumentActionsLogic] Inserting summary list at top for ${openMatches.length} open matches...`
    );
    const expanded: any[] = [];
    const config = actionPointsSettingsManager.load(e);
    openMatches.forEach((m: any) => {
      m.assignees.forEach((a: any) => {
        const entry = config.peopleConfig[a] || {};
        expanded.push({
          person: a,
          action: m.action,
          date: m.date,
          order: typeof entry.order === 'number' ? entry.order : Number.MAX_SAFE_INTEGER,
        });
      });
    });

    expanded.sort(compareActionPoints);

    body.insertParagraph(0, 'Action points').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    let lastPerson: string | null = null;
    for (let i = expanded.length - 1; i >= 0; i--) {
      const item = expanded[i];
      if (lastPerson !== null && item.person !== lastPerson) {
        body.insertParagraph(1, '');
      }
      const text = formatActionPoint(item.person, item.action, item.date);
      const listItem = body.insertListItem(1, text);
      listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
      lastPerson = item.person;
    }
    topCount = expanded.length;
    console.log(
      `[applyDocumentActionsLogic] Inserted ${topCount} summary items at top of document.`
    );
  }

  const message = `Document updates applied (${replaceCount} formatted in-place, ${topCount} added to summary).`;
  console.log(`[applyDocumentActionsLogic] Completed: ${message}`);
  return { ok: true, message, successCount: replaceCount + topCount };
}

export function savePeopleConfigFromFormLogic(formInput: Record<string, any>) {
  const validation = actionPointsSettingsManager.validate(formInput);
  if (!validation.ok) return { ok: false, message: validation.message };

  const res = actionPointsSettingsManager.save(formInput);
  if (!res.ok) return { ok: false, message: res.message };

  const config = actionPointsSettingsManager.load();
  return {
    ok: true,
    message: 'Saved',
    todoistToken: config.todoistToken,
    todoistProjectId: config.todoistProjectId,
    todoistEnabled: config.enableTodoist,
    peopleConfig: config.peopleConfig,
  };
}

export function populatePeopleConfigLogic(
  e: any,
  targetLayer: string = 'global'
): ActionPointsOperationResult {
  console.log(`[populatePeopleConfigLogic] Adding detected people to ${targetLayer} settings...`);
  const params = (e && e.parameters) || {};
  let people: string[] = [];

  if (params.peopleJson) {
    try {
      people = JSON.parse(params.peopleJson);
    } catch (err) {
      console.warn('[populatePeopleConfigLogic] Failed to parse peopleJson parameter: ', err);
    }
  }

  if (!Array.isArray(people) || people.length === 0) {
    console.log('[populatePeopleConfigLogic] No people provided; scanning active document...');
    const scanResult = scanActionPointsDocument();
    if (scanResult) {
      people = Object.keys(scanResult.foundPeople).filter((n) => n.toLowerCase() !== 'everyone');
      console.log(`[populatePeopleConfigLogic] Document scan detected ${people.length} people.`);
    }
  }

  if (!Array.isArray(people) || people.length === 0) {
    console.log('[populatePeopleConfigLogic] No detected people to add.');
    return { ok: true, message: 'No people to add.', addedCount: 0 };
  }

  const config = actionPointsSettingsManager.load(e);
  const nextPeopleConfig = { ...config.peopleConfig };
  let added = 0;

  const peopleDict: Record<string, boolean> = {};
  people.forEach((p) => {
    if (p && typeof p === 'string') peopleDict[p] = true;
  });
  const trulyNewPeople = getNewDetectedPeople(peopleDict, nextPeopleConfig);

  trulyNewPeople.forEach((name) => {
    nextPeopleConfig[name] = { aliases: [] };
    added++;
  });

  const result = actionPointsSettingsManager.save(
    { peopleConfig: nextPeopleConfig },
    targetLayer as 'global' | 'workspace',
    e
  );

  if (!result.ok) {
    console.warn(`[populatePeopleConfigLogic] Failed to save settings: ${result.message}`);
    return { ok: false, message: result.message || 'People Config is not valid JSON.' };
  }

  const msg =
    added > 0 ? `Added ${added} people to ${targetLayer} settings.` : 'No new people to add.';
  console.log(`[populatePeopleConfigLogic] Completed: ${msg}`);
  return { ok: true, message: msg, addedCount: added };
}
