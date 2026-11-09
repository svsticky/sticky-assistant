import { getTools } from '../tools/registry';
import { Tool, ToolTrigger, TriggerEvent } from './Tool';
import { getAppName, COLORS, ICON_URLS } from './branding';

type TriggerEventObject = GoogleAppsScript.Addons.EventObject & {
  parameters?: Record<string, string>;
};

function isEnabled(trigger: ToolTrigger, e: GoogleAppsScript.Addons.EventObject) {
  if (!trigger.enabled) {
    return true;
  }
  try {
    return trigger.enabled(e);
  } catch {
    return false;
  }
}

function getMatchingTriggers(event: TriggerEvent, e: GoogleAppsScript.Addons.EventObject) {
  return getTools()
    .flatMap((tool: Tool) => tool.triggers.map((trigger) => ({ tool, trigger })))
    .filter(({ trigger }) => trigger.event === event)
    .filter(({ trigger }) => isEnabled(trigger, e));
}

/**
 * Dynamically deduces the host app or Drive folder/file selection required for a tool.
 */
function getRequiredServiceDescription(tool: Tool): string {
  // Check if any trigger is Calendar homepage
  const hasCalendar = tool.triggers.some((t) => t.event === TriggerEvent.CALENDAR_HOMEPAGE);
  if (hasCalendar) return 'Google Calendar';

  // Check if any trigger is Docs homepage
  const hasDocs = tool.triggers.some((t) => t.event === TriggerEvent.DOCS_HOMEPAGE);
  if (hasDocs) return 'Google Docs';

  // Check if any trigger requires specific Drive selection
  const hasSelectedItems = tool.triggers.some((t) => t.event === TriggerEvent.ITEMS_SELECTED);
  if (hasSelectedItems) {
    if (tool.id === 'commentsExtractor') {
      return 'Google Drive (Select a PDF file)';
    }
    if (tool.id === 'committeeCreator') {
      return 'Google Drive (Select a Folder)';
    }
    return 'Google Drive (Select items)';
  }

  const hasDrive = tool.triggers.some((t) => t.event === TriggerEvent.DRIVE_HOMEPAGE);
  if (hasDrive) return 'Google Drive';

  return 'Google Workspace';
}

export function openTool(e: TriggerEventObject): GoogleAppsScript.Card_Service.Card {
  const params = e.parameters || {};
  const toolId = params.toolId;
  const eventValue = Number(params.event ?? TriggerEvent.DEFAULT_HOMEPAGE);
  const tool = getTools().find((t) => t.id === toolId);
  if (!tool) {
    return buildToolSelectorCard(e, eventValue);
  }

  const trigger = tool.triggers.find((t) => t.event === eventValue && isEnabled(t, e));
  if (!trigger) {
    return buildToolSelectorCard(e, eventValue);
  }

  return trigger.createCard(e);
}

export function buildToolSelectorCard(e: GoogleAppsScript.Addons.EventObject, event: TriggerEvent) {
  const matches = getMatchingTriggers(event, e);

  // Auto-skip selector only when targeted items are selected in Drive (targeted action)
  if (matches.length === 1 && event === TriggerEvent.ITEMS_SELECTED) {
    return matches[0].trigger.createCard(e);
  }

  const builder = CardService.newCardBuilder().setHeader(
    CardService.newCardHeader().setTitle(
      `<font color="${COLORS.PRIMARY}"><b>${getAppName()}</b></font>`
    )
  );

  const section = CardService.newCardSection().addWidget(
    CardService.newDecoratedText()
      .setStartIcon(CardService.newIconImage().setIconUrl(ICON_URLS.info))
      .setText('<b>Welcome to Sticky Assistant!</b>')
      .setBottomLabel('Select a specialized tool to get started:')
      .setWrapText(true)
  );

  // Separate active/available tools from unavailable ones in this context
  const activeToolIds = new Set(matches.map((m) => m.tool.id));
  const availableTools = matches;
  const unavailableTools = getTools().filter((tool) => !activeToolIds.has(tool.id));

  // --- SECTION 1: AVAILABLE TOOLS ---
  availableTools.forEach(({ tool, trigger }) => {
    const openAction = CardService.newAction()
      .setFunctionName('openTool')
      .setParameters({ toolId: tool.id, event: String(trigger.event) });

    section.addWidget(
      CardService.newDecoratedText()
        .setStartIcon(
          CardService.newIconImage().setIconUrl(ICON_URLS[tool.id as keyof typeof ICON_URLS] || '')
        )
        .setText(`<b>${tool.name}</b>`)
        .setBottomLabel(tool.info || 'No description available.')
        .setWrapText(true)
        .setButton(
          CardService.newTextButton()
            .setText('Open')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(openAction)
        )
    );
  });
  builder.addSection(section);

  // --- SECTION 2: OTHER TOOLS IN SUITE (UNAVAILABLE CONTEXTS) ---
  if (unavailableTools.length > 0) {
    const unavailableSection = CardService.newCardSection()
      .setHeader('Other Tools in Suite')
      .setCollapsible(true);

    unavailableTools.forEach((tool) => {
      unavailableSection.addWidget(
        CardService.newDecoratedText()
          .setStartIcon(
            CardService.newIconImage().setIconUrl(
              ICON_URLS[tool.id as keyof typeof ICON_URLS] || ''
            )
          )
          .setText(`<b>${tool.name}</b>`)
          .setBottomLabel(
            `<font color="${COLORS.MUTED}">${tool.info || 'No description available.'}</font><br><font color="${COLORS.WARNING}">⚠️ Required: ${getRequiredServiceDescription(tool)}</font>`
          )
          .setWrapText(true)
      );
    });
    builder.addSection(unavailableSection);
  }

  return builder.build();
}

/**
 * Action callback that routes the user to the correct homepage tool selector
 * based on the host application context (Calendar, Docs, Drive, or Default).
 */
export function openToolSelector(e: any): GoogleAppsScript.Card_Service.Card {
  const hostApp = e?.commonEventObject?.hostApp;
  let event = TriggerEvent.DEFAULT_HOMEPAGE;
  if (hostApp === 'CALENDAR') {
    event = TriggerEvent.CALENDAR_HOMEPAGE;
  } else if (hostApp === 'DOCS') {
    event = TriggerEvent.DOCS_HOMEPAGE;
  } else if (hostApp === 'DRIVE') {
    event = TriggerEvent.DRIVE_HOMEPAGE;
  }
  return buildToolSelectorCard(e, event);
}
