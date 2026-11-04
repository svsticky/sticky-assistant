export interface PeopleConfigEntry {
  aliases?: string[];
  todoist_id?: string;
  todoist_section?: string;
  order?: number;
}

export interface ActionPointsConfig {
  todoistToken: string;
  todoistProjectId: string;
  enableTodoist: boolean;
  peopleConfig: Record<string, PeopleConfigEntry>;
}

export function validateActionPointsConfig(formInput: Record<string, any>): {
  ok: boolean;
  message?: string;
} {
  if (formInput.peopleConfig !== undefined) {
    try {
      const parsed = JSON.parse(String(formInput.peopleConfig || '{}'));
      return validatePeopleConfig(parsed);
    } catch {
      return { ok: false, message: 'People Config must be valid JSON.' };
    }
  }
  // Basic validation for token/project ID lengths (optional)
  if (formInput.todoistToken !== undefined && String(formInput.todoistToken).length > 1024) {
    return { ok: false, message: 'Todoist token is too long.' };
  }
  return { ok: true };
}

export function parsePeopleConfig(raw: string): Record<string, PeopleConfigEntry> {
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return validatePeopleConfig(parsed).ok ? parsed : {};
  } catch {
    return {};
  }
}

export function formatPeopleConfig(config: Record<string, PeopleConfigEntry>): string {
  return JSON.stringify(config || {}, null, 2);
}

export function validatePeopleConfig(config: any): { ok: boolean; message?: string } {
  if (config === null || typeof config !== 'object' || Array.isArray(config)) {
    return { ok: false, message: 'People Config must be a JSON object.' };
  }

  const keys = Object.keys(config);
  for (let i = 0; i < keys.length; i++) {
    const personName = keys[i];
    const entry = config[personName];
    if (entry === null || typeof entry !== 'object' || Array.isArray(entry)) {
      return { ok: false, message: `Entry for "${personName}" must be an object.` };
    }

    if (entry.aliases !== undefined) {
      if (!Array.isArray(entry.aliases)) {
        return { ok: false, message: `Aliases for "${personName}" must be an array.` };
      }
      for (let j = 0; j < entry.aliases.length; j++) {
        if (typeof entry.aliases[j] !== 'string') {
          return { ok: false, message: `Aliases for "${personName}" must be strings.` };
        }
      }
    }

    if (entry.todoist_id !== undefined && typeof entry.todoist_id !== 'string') {
      return { ok: false, message: `todoist_id for "${personName}" must be a string.` };
    }

    if (entry.todoist_section !== undefined && typeof entry.todoist_section !== 'string') {
      return { ok: false, message: `todoist_section for "${personName}" must be a string.` };
    }

    if (entry.order !== undefined && typeof entry.order !== 'number') {
      return { ok: false, message: `order for "${personName}" must be a number.` };
    }
  }

  return { ok: true };
}
