import { describe, it, expect, beforeEach, vi } from 'vitest';

vi.hoisted(() => {
  globalThis.CardService = {
    Icon: {
      STORE: 'STORE_ICON',
    },
  } as any;
});

import {
  applyDocumentActionsLogic,
  formatTodoistExportTask,
  getNewDetectedPeople,
  populatePeopleConfigLogic,
} from './editing';
import { compareActionPoints, parseMeetingDate, scanActionPointsDocument } from './scanning';

// Mock properties/settings
const mockUserProperties: Record<string, string> = {};
globalThis.PropertiesService = {
  getUserProperties: () => ({
    getProperty: (key: string) => mockUserProperties[key] ?? null,
    setProperty: (key: string, value: string) => {
      mockUserProperties[key] = value;
    },
  }),
} as any;

class MockElement {
  type: any;
  text: string;
  isBoldValue: boolean;
  strikethrough: boolean = false;

  constructor(type: any, text: string, isBold: boolean = false) {
    this.type = type;
    this.text = text;
    this.isBoldValue = isBold;
  }

  getType() {
    return this.type;
  }

  getText() {
    return this.text;
  }

  editAsText() {
    return this;
  }

  asText() {
    return this;
  }

  isBold(_index: number) {
    return this.isBoldValue;
  }

  setStrikethrough(_start: number, _end: number, strikethrough: boolean) {
    this.strikethrough = strikethrough;
    return this;
  }
}

class MockParagraph extends MockElement {
  heading: any;
  setHeading(heading: any) {
    this.heading = heading;
    return this;
  }
}

class MockListItem extends MockElement {
  glyphType: any;
  nestingLevel: number = 0;
  getNestingLevel() {
    return this.nestingLevel;
  }
  setGlyphType(glyphType: any) {
    this.glyphType = glyphType;
    return this;
  }
}

let mockChildren: MockElement[] = [];

const mockBody = {
  getChildren: () => mockChildren,
  getNumChildren: () => mockChildren.length,
  getChild: (index: number) => mockChildren[index],
  insertParagraph: (index: number, text: string) => {
    const p = new MockParagraph('PARAGRAPH', text);
    mockChildren.splice(index, 0, p);
    return p;
  },
  insertListItem: (index: number, text: string) => {
    const li = new MockListItem('LIST_ITEM', text);
    mockChildren.splice(index, 0, li);
    return li;
  },
  findText: (pattern: string) => {
    for (const child of mockChildren) {
      if (child.text.indexOf(pattern) !== -1) {
        return {
          getElement: () => child,
          getStartOffset: () => child.text.indexOf(pattern),
          getEndOffsetInclusive: () => child.text.indexOf(pattern) + pattern.length - 1,
        };
      }
    }
    return null;
  },
};

globalThis.DocumentApp = {
  ParagraphHeading: {
    HEADING2: 'HEADING2',
  },
  GlyphType: {
    BULLET: 'BULLET',
  },
  ElementType: {
    PARAGRAPH: 'PARAGRAPH',
    LIST_ITEM: 'LIST_ITEM',
  },
  getActiveDocument: () => ({
    getId: () => 'mock-doc-id',
    getName: () => 'Mock Document',
    getBody: () => mockBody,
  }),
} as any;

globalThis.DriveApp = {
  getFileById: (_id: string) => ({
    getParents: () => ({
      hasNext: () => false,
      next: () => null,
    }),
  }),
} as any;

describe('applyDocumentActionsLogic', () => {
  beforeEach(() => {
    mockChildren = [];
    Object.keys(mockUserProperties).forEach((k) => delete mockUserProperties[k]);
  });

  it('fails gracefully when no document is active', () => {
    const oldGetActiveDocument = globalThis.DocumentApp.getActiveDocument;
    globalThis.DocumentApp.getActiveDocument = () => null as any;

    try {
      const res = applyDocumentActionsLogic({});
      expect(res.ok).toBe(false);
      expect(res.message).toBe('No active Google Doc found.');
    } finally {
      globalThis.DocumentApp.getActiveDocument = oldGetActiveDocument;
    }
  });

  it('should parse matchesJson and execute addToTop and replaceInPlace correctly', () => {
    // Add an initial paragraph to mock document
    const initialParagraph = new MockParagraph('PARAGRAPH', 'AP Ruben: Finished Task', false);
    mockChildren.push(initialParagraph);

    // Mock settings
    mockUserProperties['actionPointsExtractor__peopleConfig'] = JSON.stringify({
      Ruben: { order: 1, aliases: [] },
    });

    const event = {
      parameters: {
        matchesJson: JSON.stringify({
          openMatches: [
            {
              originalName: 'Ruben',
              assignees: ['Ruben'],
              action: 'Task 1',
              date: '2026-06-15',
              isBold: true,
              location: {
                childIndex: 0,
                matchIndex: 0,
                matchLength: 22,
              },
              originalText: 'AP Ruben: Task 1',
            },
          ],
          completedMatches: [
            {
              originalName: 'Ruben',
              assignees: ['Ruben'],
              action: 'Finished Task',
              date: null,
              isBold: false,
              location: {
                childIndex: 0,
                matchIndex: 0,
                matchLength: 23,
              },
              originalText: 'AP Ruben: Finished Task',
            },
          ],
        }),
      },
      formInput: {
        addToTopAction: 'addToTop',
        replaceInPlaceAction: 'replaceInPlace',
      },
    };

    const res = applyDocumentActionsLogic(event);
    expect(res.ok).toBe(true);

    // Verify that strikethrough was set to true on the completed task in the document
    expect(initialParagraph.strikethrough).toBe(true);

    // Verify that "Action points" header and bullet list item were inserted at the top (indices 0 and 1)
    expect(mockChildren[0].text).toBe('Action points');
    expect(mockChildren[0].getType()).toBe('PARAGRAPH');
    expect((mockChildren[0] as MockParagraph).heading).toBe('HEADING2');

    expect(mockChildren[1].text).toBe('AP Ruben: Task 1 [2026-06-15]');
    expect(mockChildren[1].getType()).toBe('LIST_ITEM');
    expect((mockChildren[1] as MockListItem).glyphType).toBe('BULLET');
  });

  it('scans document directly when matchesJson is not provided in parameters', () => {
    // An unbolded completed task and a bold open task
    const completedParagraph = new MockParagraph('PARAGRAPH', 'AP Aron: Completed task', false);
    const openParagraph = new MockParagraph('PARAGRAPH', 'AP Aron: Open task', true);
    mockChildren.push(completedParagraph);
    mockChildren.push(openParagraph);

    const event = {
      parameters: {}, // No matchesJson!
      formInput: {
        addToTopAction: 'addToTop',
        replaceInPlaceAction: 'replaceInPlace',
      },
    };

    const res = applyDocumentActionsLogic(event);
    expect(res.ok).toBe(true);
    expect(completedParagraph.strikethrough).toBe(true);

    // Verify summary inserted at top
    expect(mockChildren[0].text).toBe('Action points');
    expect(mockChildren[1].text).toBe('AP Aron: Open task');
  });

  it('returns graceful message when no checkbox actions are selected', () => {
    const event = {
      parameters: {},
      formInput: {},
    };

    const res = applyDocumentActionsLogic(event);
    expect(res.ok).toBe(true);
    expect(res.message).toBe('No actions selected. Please check at least one box.');
  });
});

describe('formatTodoistExportTask', () => {
  it('formats correctly with date, person (todoist_id), and section configured', () => {
    const task = {
      person: 'Ruben',
      action: 'Write code',
      date: '2026-07-01',
      order: 1,
    };
    const peopleConfig = {
      Ruben: {
        todoist_id: 'RubenBroere',
        todoist_section: 'Development',
      },
    };
    const result = formatTodoistExportTask(task, peopleConfig);
    expect(result).toBe('[AP] Write code /Development 2026-07-01 +RubenBroere');
  });

  it('omits date when date is null and no defaultDueDate is provided', () => {
    const task = {
      person: 'Bob',
      action: 'Test application',
      date: null,
      order: 3,
    };
    const peopleConfig = {
      Bob: {
        todoist_id: 'BobTester',
        todoist_section: 'QA',
      },
    };
    const result = formatTodoistExportTask(task, peopleConfig);
    expect(result).toBe('[AP] Test application /QA +BobTester');
  });

  it('preserves date in action text and moves due date and assignee to the end', () => {
    const task = {
      person: 'Ruben',
      action: 'Prepare proposal for 2026-11-01 meeting',
      date: '2026-11-12',
      order: 1,
    };
    const peopleConfig = {
      Ruben: {
        todoist_id: 'RubenBroere',
      },
    };
    const result = formatTodoistExportTask(task, peopleConfig);
    expect(result).toBe('[AP] Prepare proposal for 2026-11-01 meeting 2026-11-12 +RubenBroere');
  });

  it('falls back to task.person when not in configuration', () => {
    const task = {
      person: 'Charlie',
      action: 'Deploy app',
      date: '2026-07-03',
      order: 4,
    };
    const peopleConfig = {};
    const result = formatTodoistExportTask(task, peopleConfig);
    expect(result).toBe('[AP] Deploy app 2026-07-03 +Charlie');
  });

  it('falls back to task.person when in configuration but without todoist_id', () => {
    const task = {
      person: 'Diana',
      action: 'Write docs',
      date: '2026-07-04',
      order: 5,
    };
    const peopleConfig = {
      Diana: {
        // no todoist_id
      },
    };
    const result = formatTodoistExportTask(task, peopleConfig);
    expect(result).toBe('[AP] Write docs 2026-07-04 +Diana');
  });

  it('uses defaultDueDate when task date is null', () => {
    const task = {
      person: 'Alice',
      action: 'Run task',
      date: null,
      order: 1,
    };
    const peopleConfig = {
      Alice: {
        todoist_id: 'AliceA',
      },
    };
    const result = formatTodoistExportTask(task, peopleConfig, '2026-07-15');
    expect(result).toBe('[AP] Run task 2026-07-15 +AliceA');
  });
});

describe('getNewDetectedPeople', () => {
  it('returns empty array when all detected people are already in peopleConfig', () => {
    const found = { Ruben: true, Aron: true, Toine: true };
    const peopleConfig = {
      Ruben: { aliases: ['ruben b'] },
      Aron: {},
      Toine: {},
    };
    expect(getNewDetectedPeople(found, peopleConfig)).toEqual([]);
  });

  it('matches existing people case-insensitively and via aliases', () => {
    const found = { ruben: true, chef: true };
    const peopleConfig = {
      Ruben: { aliases: ['Chef', 'ruben b'] },
    };
    expect(getNewDetectedPeople(found, peopleConfig)).toEqual([]);
  });

  it('ignores everyone and iedereen', () => {
    const found = { everyone: true, Iedereen: true };
    const peopleConfig = {};
    expect(getNewDetectedPeople(found, peopleConfig)).toEqual([]);
  });

  it('returns only new unconfigured people', () => {
    const found = { Ruben: true, Kees: true, Jan: true, everyone: true };
    const peopleConfig = {
      Ruben: {},
    };
    expect(getNewDetectedPeople(found, peopleConfig)).toEqual(['Kees', 'Jan']);
  });
});

describe('compareActionPoints', () => {
  it('sorts by order ascending', () => {
    const tasks = [
      { person: 'Charlie', order: 3 },
      { person: 'Alice', order: 1 },
      { person: 'Bob', order: 2 },
    ];
    tasks.sort(compareActionPoints);
    expect(tasks.map((t) => t.person)).toEqual(['Alice', 'Bob', 'Charlie']);
  });

  it('puts null and undefined orders after defined numerical orders', () => {
    const tasks = [
      { person: 'NoOrder1', order: null },
      { person: 'First', order: 1 },
      { person: 'NoOrder2' },
      { person: 'Second', order: 2 },
    ];
    tasks.sort(compareActionPoints);
    expect(tasks.map((t) => t.person)).toEqual(['First', 'Second', 'NoOrder1', 'NoOrder2']);
  });

  it('breaks ties alphabetically by person name', () => {
    const tasks = [
      { person: 'Zach', order: 1 },
      { person: 'Alice', order: 1 },
      { person: 'Bob', order: null },
      { person: 'Aaron', order: null },
    ];
    tasks.sort(compareActionPoints);
    expect(tasks.map((t) => t.person)).toEqual(['Alice', 'Zach', 'Aaron', 'Bob']);
  });
});

describe('populatePeopleConfigLogic', () => {
  beforeEach(() => {
    mockUserProperties['actionPointsExtractor__peopleConfig'] = JSON.stringify({
      Ruben: { aliases: [] },
    });
  });

  it('adds newly detected people to configuration', () => {
    const event = {
      parameters: {
        peopleJson: JSON.stringify(['Kees', 'Jan', 'Ruben']),
      },
    };
    const res = populatePeopleConfigLogic(event, 'global');
    expect(res.ok).toBe(true);
    expect(res.addedCount).toBe(2);

    const savedConfig = JSON.parse(
      mockUserProperties['actionPointsExtractor__peopleConfig'] || '{}'
    );
    expect(savedConfig.Kees).toBeDefined();
    expect(savedConfig.Jan).toBeDefined();
    expect(savedConfig.Ruben).toBeDefined();
  });

  it('returns gracefully when no new people are found', () => {
    const event = {
      parameters: {
        peopleJson: JSON.stringify(['Ruben', 'everyone']),
      },
    };
    const res = populatePeopleConfigLogic(event, 'global');
    expect(res.ok).toBe(true);
    expect(res.addedCount).toBe(0);
  });
});

describe('parseMeetingDate', () => {
  it('parses YYYY-MM-DD format', () => {
    expect(parseMeetingDate('Volgende meeting is 2026-07-15.')).toBe('2026-07-15');
  });

  it('parses YYYY/MM/DD format', () => {
    expect(parseMeetingDate('Volgende meeting is 2026/07/15.')).toBe('2026-07-15');
  });

  it('parses DD-MM-YYYY format', () => {
    expect(parseMeetingDate('Volgende meeting is 15-07-2026.')).toBe('2026-07-15');
  });

  it('parses DD/MM/YYYY format', () => {
    expect(parseMeetingDate('Volgende meeting is 15/07/2026.')).toBe('2026-07-15');
  });

  it('returns null on invalid formats', () => {
    expect(parseMeetingDate('No meeting date specified')).toBeNull();
    expect(parseMeetingDate('dinsdag 7 juli 2026')).toBeNull();
  });
});

describe('scanActionPointsDocument next meeting date', () => {
  beforeEach(() => {
    mockChildren = [];
  });

  it('finds nested meeting date under Volgende vergadering list item', () => {
    const heading = new MockListItem('LIST_ITEM', 'Volgende vergadering');
    heading.nestingLevel = 0;
    const dateItem = new MockListItem('LIST_ITEM', 'Volgende vergadering op 07-07-2026');
    dateItem.nestingLevel = 1;

    mockChildren.push(heading, dateItem);

    const result = scanActionPointsDocument();
    expect(result?.nextMeetingDate).toBe('2026-07-07');
  });

  it('finds nested meeting date under numbered 22. Volgende vergadering (screenshot scenario)', () => {
    const heading = new MockParagraph('PARAGRAPH', '22. Volgende vergadering');
    const dateItem = new MockListItem('LIST_ITEM', 'a. 2026-11-12');
    const agendaItem = new MockListItem('LIST_ITEM', 'b. Agendapunten');
    const closingItem = new MockParagraph('PARAGRAPH', '23. Sluiting');

    mockChildren.push(heading, dateItem, agendaItem, closingItem);

    const result = scanActionPointsDocument();
    expect(result?.nextMeetingDate).toBe('2026-11-12');
  });

  it('finds meeting date on the same line as header', () => {
    const heading = new MockParagraph('PARAGRAPH', '22. Volgende vergadering: 2026-11-12');
    mockChildren.push(heading);

    const result = scanActionPointsDocument();
    expect(result?.nextMeetingDate).toBe('2026-11-12');
  });

  it('finds nested meeting date under Next meeting paragraph', () => {
    const heading = new MockParagraph('PARAGRAPH', 'Next meeting:');
    const dateItem = new MockListItem('LIST_ITEM', '15-09-2026 - online');
    dateItem.nestingLevel = 0;

    mockChildren.push(heading, dateItem);

    const result = scanActionPointsDocument();
    expect(result?.nextMeetingDate).toBe('2026-09-15');
  });

  it('continues searching document if an earlier candidate did not contain a meeting date', () => {
    // Earlier candidate that was just an agenda item or lacked a date
    const falseHeading = new MockListItem('LIST_ITEM', 'Volgende vergadering');
    const falseChild = new MockListItem('LIST_ITEM', 'Ruben: Helix lustrum stukje [5m]');
    // Real meeting section later in the document
    const realHeading = new MockParagraph('PARAGRAPH', '22. Volgende vergadering');
    const realDate = new MockListItem('LIST_ITEM', 'a. 2026-11-12');

    mockChildren.push(falseHeading, falseChild, realHeading, realDate);

    const result = scanActionPointsDocument();
    expect(result?.nextMeetingDate).toBe('2026-11-12');
  });

  it('stops search if sibling or parent element is encountered', () => {
    const heading = new MockListItem('LIST_ITEM', 'Volgende vergadering');
    heading.nestingLevel = 1;

    const anotherHeading = new MockListItem('LIST_ITEM', 'Other topic');
    anotherHeading.nestingLevel = 1;

    const dateItem = new MockListItem('LIST_ITEM', '07-07-2026');
    dateItem.nestingLevel = 2; // indented under "Other topic"

    mockChildren.push(heading, anotherHeading, dateItem);

    const result = scanActionPointsDocument();
    expect(result?.nextMeetingDate).toBeNull();
  });
});
