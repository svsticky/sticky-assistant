import { actionPointsSettingsManager } from './settings';
import { PeopleConfigEntry } from './config';
import { ActionPointTask } from './scanning';

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

function getTodayDateString(): string {
  if (typeof Utilities !== 'undefined' && typeof Session !== 'undefined') {
    try {
      return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch {
      // fallback
    }
  }
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const date = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${date}`;
}

export function formatTodoistExportTask(
  task: ActionPointTask,
  peopleConfig: Record<string, PeopleConfigEntry>,
  defaultDueDate?: string | null
): string {
  const parts: string[] = ['[AP]'];

  const date = task.date || defaultDueDate || getTodayDateString();
  parts.push(date);

  const entry = peopleConfig[task.person];
  const todoistId = entry?.todoist_id || task.person;
  if (todoistId) {
    parts.push(`+${todoistId}`);
  }

  if (entry?.todoist_section) {
    parts.push(`/${entry.todoist_section}`);
  }

  parts.push(task.action);

  return parts.join(' ');
}

export function sendToTodoistLogic(e: any) {
  const tasks = Array.isArray(e)
    ? e
    : JSON.parse((e && e.parameters && e.parameters.tasksJson) || '[]');
  const config = actionPointsSettingsManager.load();
  if (!config.todoistToken || !config.todoistProjectId) {
    return { ok: false, message: 'Please add your Todoist Token and Project ID in Settings.' };
  }

  let successCount = 0;
  const url = 'https://api.todoist.com/rest/v2/tasks';

  tasks.forEach((task: any) => {
    const content = `${task.person}: ${task.action}`;
    const payloadObj: any = { content, project_id: config.todoistProjectId };

    if (task.date) {
      const parts = task.date.split('-');
      if (parts.length === 3) {
        payloadObj.due_date = `${parts[2]}-${parts[1]}-${parts[0]}`;
      }
    }

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
    }
  });

  return { ok: true, message: `Synced ${successCount} tasks to Todoist.`, successCount };
}

export function applyDocumentActionsLogic(e: any) {
  const parsed = Array.isArray(e)
    ? e
    : JSON.parse(
        (e && e.parameters && (e.parameters.matchesJson || e.parameters.resultJson)) || '{}'
      );

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

  const doc = DocumentApp.getActiveDocument();
  if (!doc) {
    return { ok: false, message: 'No active Google Doc found.' };
  }

  const body = doc.getBody();

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
  if (replaceInPlace && parsed.completedMatches && parsed.completedMatches.length > 0) {
    try {
      doReplace(parsed.completedMatches, true);
    } catch {
      // Allow addToTop to proceed even if a replacement fails
    }
  }

  if (addToTop) {
    const openMatches = parsed.openMatches || [];
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

    expanded.sort((x, y) => {
      if (x.order !== y.order) return x.order - y.order;
      return x.person.localeCompare(y.person);
    });

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
  }

  return { ok: true, message: 'Document updates applied.' };
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

export function populatePeopleConfigLogic(e: any, targetLayer: string = 'global') {
  const params = e.parameters || {};
  const peopleJson = params.peopleJson || '[]';
  let people: any[];
  try {
    people = JSON.parse(peopleJson);
  } catch {
    return { ok: false, message: 'Could not read people list.' };
  }

  if (!Array.isArray(people) || people.length === 0) {
    return { ok: true, message: 'No people to add.', addedCount: 0 };
  }

  const config = actionPointsSettingsManager.load(e);
  const nextPeopleConfig = { ...config.peopleConfig };
  let added = 0;

  people.forEach((name) => {
    if (!name || typeof name !== 'string') return;
    if (name.toLowerCase() === 'everyone') return;
    if (!nextPeopleConfig[name]) {
      nextPeopleConfig[name] = { aliases: [] };
      added++;
    }
  });

  const result = actionPointsSettingsManager.save(
    { peopleConfig: nextPeopleConfig },
    targetLayer as 'global' | 'workspace',
    e
  );

  if (!result.ok) {
    return { ok: false, message: result.message || 'People Config is not valid JSON.' };
  }

  const msg =
    added > 0 ? `Added ${added} people to ${targetLayer} settings.` : 'No new people to add.';
  return { ok: true, message: msg, addedCount: added };
}
