import { getActiveParentFolder } from '../../core/settingsStore';
import { ActionPointsScanResult, scanActionPointsDocument } from './scanning';
import { actionPointsSettingsManager, ACTION_POINTS_SETTINGS } from './settings';
import {
  applyDocumentActionsLogic,
  sendToTodoistLogic,
  populatePeopleConfigLogic,
  formatTodoistExportTask,
} from './editing';
import { buildToolCard, buildToolFooter } from '../../core/cardTemplate';
import { COLORS, ICON_URLS } from '../../core/branding';

function createActionPointsBuilder() {
  const toolMeta = {
    id: 'actionPointsExtractor',
    name: 'Action Points',
    icon: CardService.Icon.STORE,
    settings: ACTION_POINTS_SETTINGS,
    triggers: [],
  };
  return buildToolCard(toolMeta, 'Scan and sync action items.');
}

export function createActionPointsHomepage(): GoogleAppsScript.Card_Service.Card {
  const builder = createActionPointsBuilder();
  const section = CardService.newCardSection();

  const scanAction = CardService.newAction().setFunctionName('scanDocument');
  section.addWidget(
    CardService.newTextButton()
      .setText('Scan Document')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(scanAction)
  );

  builder.addSection(section);
  builder.addSection(buildToolFooter('actionPointsExtractor', true));
  return builder.build();
}

export function buildActionPointsScanResultsCard(
  scanResult: ActionPointsScanResult
): GoogleAppsScript.Card_Service.Card {
  const builder = createActionPointsBuilder();
  builder.setHeader(CardService.newCardHeader().setTitle('Scan Results'));

  // --- COMPLETION SUMMARY (PROGRESS BAR) ---
  const total = scanResult.openTasks.length + scanResult.completedTasks.length;

  const summarySection = CardService.newCardSection().setHeader('Completion Summary');
  summarySection.addWidget(
    CardService.newDecoratedText()
      .setStartIcon(CardService.newIconImage().setIconUrl(ICON_URLS.success))
      .setText('<b>AP Progress</b>')
      .setBottomLabel(`${scanResult.completedTasks.length} / ${total} Tasks Completed`)
      .setWrapText(true)
  );
  builder.addSection(summarySection);

  // --- OPEN ACTION POINTS ---
  const openSection = CardService.newCardSection()
    .setHeader('Open Action Points')
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(3);

  if (scanResult.openTasks.length === 0) {
    openSection.addWidget(
      CardService.newDecoratedText()
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.STAR))
        .setText(`<font color="${COLORS.MUTED}"><i>No open action points found.</i></font>`)
    );
  } else {
    scanResult.openTasks.forEach((task) => {
      const showAction = CardService.newAction()
        .setFunctionName('jumpToTask')
        .setParameters({ childIndex: String(task.childIndex ?? 0) });

      openSection.addWidget(
        CardService.newDecoratedText()
          .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CLOCK))
          .setText(`<b>${task.person}</b>: ${task.action}`)
          .setBottomLabel(task.date ? `Due: ${task.date}` : 'No due date')
          .setWrapText(true)
          .setButton(CardService.newTextButton().setText('Show').setOnClickAction(showAction))
      );
    });
  }
  builder.addSection(openSection);

  // --- TODOIST PASTE EXPORT ---
  if (scanResult.openTasks.length > 0) {
    const todoistExportSection = CardService.newCardSection()
      .setHeader('Todoist Paste Export')
      .setCollapsible(true);

    const config = actionPointsSettingsManager.load();
    const peopleConfig = config.peopleConfig || {};

    const sortedTasks = [...scanResult.openTasks].sort((x, y) => {
      const xOrder = x.order !== null && x.order !== undefined ? x.order : Number.MAX_SAFE_INTEGER;
      const yOrder = y.order !== null && y.order !== undefined ? y.order : Number.MAX_SAFE_INTEGER;
      if (xOrder !== yOrder) return xOrder - yOrder;
      return x.person.localeCompare(y.person);
    });

    const exportText = sortedTasks
      .map((task) => formatTodoistExportTask(task, peopleConfig, scanResult.nextMeetingDate))
      .join('\n');

    todoistExportSection.addWidget(
      CardService.newTextInput()
        .setFieldName('todoistExportText')
        .setTitle('Copy-paste tasks to Todoist')
        .setMultiline(true)
        .setValue(exportText)
    );
    builder.addSection(todoistExportSection);
  }

  // --- COMPLETED ACTION POINTS ---
  const completedSection = CardService.newCardSection()
    .setHeader('Completed Action Points')
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(0);

  if (scanResult.completedTasks.length === 0) {
    completedSection.addWidget(
      CardService.newDecoratedText()
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.CONFIRMATION_NUMBER_ICON))
        .setText(`<font color="${COLORS.MUTED}"><i>No completed action points found.</i></font>`)
    );
  } else {
    scanResult.completedTasks.forEach((task) => {
      const showAction = CardService.newAction()
        .setFunctionName('jumpToTask')
        .setParameters({ childIndex: String(task.childIndex ?? 0) });

      completedSection.addWidget(
        CardService.newDecoratedText()
          .setStartIcon(
            CardService.newIconImage().setIcon(CardService.Icon.CONFIRMATION_NUMBER_ICON)
          )
          .setText(
            `<font color="${COLORS.MUTED}"><s><b>${task.person}</b>: ${task.action}</s></font>`
          )
          .setBottomLabel(task.date ? `Completed • Due: ${task.date}` : 'Completed')
          .setWrapText(true)
          .setButton(CardService.newTextButton().setText('Show').setOnClickAction(showAction))
      );
    });
  }
  builder.addSection(completedSection);

  // --- ACTIONS SECTION ---
  const actionSection = CardService.newCardSection().setHeader('Actions');
  const config = actionPointsSettingsManager.load();

  if (config.enableTodoist && scanResult.openTasks.length > 0) {
    actionSection.addWidget(
      CardService.newTextButton()
        .setText('Sync Open Tasks to Todoist')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName('sendToTodoist')
            .setParameters({ tasksJson: JSON.stringify(scanResult.openTasks) })
        )
    );
  } else if (config.enableTodoist) {
    actionSection.addWidget(
      CardService.newDecoratedText()
        .setStartIcon(CardService.newIconImage().setIcon(CardService.Icon.STAR))
        .setText(`<font color="${COLORS.MUTED}"><i>No open action points to sync.</i></font>`)
    );
  }

  if (scanResult.openMatches.length > 0 || scanResult.completedMatches.length > 0) {
    actionSection.addWidget(
      CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setFieldName('addToTopAction')
        .addItem('Insert summary list at top', 'addToTop', false)
    );
    actionSection.addWidget(
      CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.CHECK_BOX)
        .setFieldName('replaceInPlaceAction')
        .addItem('Format completed in-place', 'replaceInPlace', false)
    );

    actionSection.addWidget(
      CardService.newTextButton()
        .setText('Apply Changes')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName('applyDocumentActions')
            .setParameters({
              matchesJson: JSON.stringify({
                openMatches: scanResult.openMatches,
                completedMatches: scanResult.completedMatches,
              }),
            })
        )
    );
  }

  const foundPeopleList = Object.keys(scanResult.foundPeople).filter(
    (name) => name.toLowerCase() !== 'everyone'
  );
  if (foundPeopleList.length > 0) {
    actionSection.addWidget(
      CardService.newTextButton()
        .setText('Add Detected People to Settings')
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName('buildAddPeopleLayerCard')
            .setParameters({ peopleJson: JSON.stringify(foundPeopleList) })
        )
    );
  }

  // Back button returning to Action Points Homepage using generic openTool action
  const backAction = CardService.newAction()
    .setFunctionName('openTool')
    .setParameters({ toolId: 'actionPointsExtractor' });

  actionSection.addWidget(CardService.newTextButton().setText('Back').setOnClickAction(backAction));

  builder.addSection(actionSection);
  return builder.build();
}

/**
 * Custom layout that lets the user choose whether to save newly detected assignees
 * to the Global configuration (User properties) or local Workspace (sticky-assistant.json) layer.
 */
export function buildAddPeopleLayerCard(e: {
  parameters?: Record<string, string>;
}): GoogleAppsScript.Card_Service.Card {
  const params = e.parameters || {};
  const peopleJson = params.peopleJson || '[]';

  const builder = CardService.newCardBuilder();
  builder.setHeader(
    CardService.newCardHeader()
      .setTitle(`<font color="${COLORS.PRIMARY}"><b>Add People to Settings</b></font>`)
      .setImageUrl(ICON_URLS.actionPointsExtractor)
  );

  const parentFolder = getActiveParentFolder(e);

  const section = CardService.newCardSection().setHeader('Add Options');
  section.addWidget(
    CardService.newDecoratedText()
      .setStartIcon(CardService.newIconImage().setIconUrl(ICON_URLS.info))
      .setText('<b>Add Detected People</b>')
      .setBottomLabel(
        'Add the newly detected assignees to either your Workspace folder config or Global account settings.'
      )
      .setWrapText(true)
  );

  // Workspace save button (primary action if folder is accessible)
  if (parentFolder) {
    const addWorkspaceAction = CardService.newAction()
      .setFunctionName('populatePeopleConfig')
      .setParameters({ peopleJson, targetLayer: 'workspace' });

    section.addWidget(
      CardService.newTextButton()
        .setText('Add to Workspace (This Folder)')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(addWorkspaceAction)
    );
  }

  // Global save button (secondary action)
  const addGlobalAction = CardService.newAction()
    .setFunctionName('populatePeopleConfig')
    .setParameters({ peopleJson, targetLayer: 'global' });

  section.addWidget(
    CardService.newTextButton().setText('Add to Global (Account)').setOnClickAction(addGlobalAction)
  );

  // Back action returning to home
  const backAction = CardService.newAction()
    .setFunctionName('openTool')
    .setParameters({ toolId: 'actionPointsExtractor' });

  section.addWidget(CardService.newTextButton().setText('Cancel').setOnClickAction(backAction));

  builder.addSection(section);
  return builder.build();
}

/**
 * Focuses/Scrolls Google Docs editor range to select and highlight the paragraph of an AP task.
 */
export function jumpToTask(e: { parameters?: Record<string, string> }) {
  const params = e.parameters || {};
  const childIndex = Number(params.childIndex || 0);

  try {
    const doc = DocumentApp.getActiveDocument();
    if (!doc) throw new Error('No active document found.');

    const body = doc.getBody();
    const child = body.getChild(childIndex);

    const rangeBuilder = doc.newRange();
    rangeBuilder.addElement(child);
    doc.setSelection(rangeBuilder.build());

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Task selected & scrolled into view'))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText('Could not locate element: ' + String(err))
      )
      .build();
  }
}

export function scanDocument() {
  const scanResult = scanActionPointsDocument();
  if (!scanResult) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('No Document Found'))
      .addSection(
        CardService.newCardSection().addWidget(
          CardService.newDecoratedText()
            .setStartIcon(CardService.newIconImage().setIconUrl(ICON_URLS.error))
            .setText(`<font color="${COLORS.ERROR}"><b>No Document Active</b></font>`)
            .setBottomLabel('Please open a Google Docs document and try again.')
        )
      )
      .build();
  }

  if (
    scanResult.openTasks.length === 0 &&
    scanResult.completedTasks.length === 0 &&
    scanResult.openMatches.length === 0 &&
    scanResult.completedMatches.length === 0
  ) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('No Action Points'))
      .addSection(
        CardService.newCardSection().addWidget(
          CardService.newDecoratedText()
            .setStartIcon(CardService.newIconImage().setIconUrl(ICON_URLS.info))
            .setText(`<font color="${COLORS.MUTED}"><b>No Action Points Found</b></font>`)
            .setBottomLabel('No open or completed action points were found in this document.')
        )
      )
      .build();
  }

  return buildActionPointsScanResultsCard(scanResult);
}

export function sendToTodoist(e: GoogleAppsScript.Addons.EventObject) {
  const result = sendToTodoistLogic(e);
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(result.message))
    .build();
}

export function applyDocumentActions(e: GoogleAppsScript.Addons.EventObject) {
  const result = applyDocumentActionsLogic(e);
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(result.message))
    .build();
}

export function populatePeopleConfig(e: any) {
  const params = e.parameters || {};
  const formInput = e.formInput || {};
  const targetLayer = params.targetLayer || formInput.targetLayer || 'global';
  const result = populatePeopleConfigLogic(e, targetLayer);
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(result.message))
    .build();
}
