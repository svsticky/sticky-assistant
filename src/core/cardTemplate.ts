import { Tool } from './Tool';
import { COLORS, ICON_URLS } from './branding';
import { APP_VERSION } from './version';

type ToolCardMeta = Pick<Tool, 'id' | 'name' | 'icon' | 'settings'>;

/**
 * Builds a standardized, beautiful base card for tools, including
 * a modern header with custom logo, and a wrapped description section.
 */
export function buildToolCard(
  tool: ToolCardMeta,
  info?: string
): GoogleAppsScript.Card_Service.CardBuilder {
  const builder = CardService.newCardBuilder();

  // Get custom remote SVG icon URL or fallback to logoUrl
  const logoUrl = ICON_URLS[tool.id as keyof typeof ICON_URLS] || '';

  // Polished header with tool name, app logo, and style accents
  const header = CardService.newCardHeader()
    .setTitle(`<font color="${COLORS.PRIMARY}"><b>${tool.name}</b></font>`)
    .setImageStyle(CardService.ImageStyle.CIRCLE);

  if (logoUrl) {
    header.setImageUrl(logoUrl);
  }
  builder.setHeader(header);

  if (info) {
    const infoSection = CardService.newCardSection().setHeader('About this Tool');
    infoSection.addWidget(
      CardService.newDecoratedText()
        .setStartIcon(CardService.newIconImage().setIconUrl(ICON_URLS.info))
        .setText(`<font color="${COLORS.MUTED}">${info}</font>`)
        .setBottomLabel(`Version ${APP_VERSION}`)
        .setWrapText(true)
    );
    builder.addSection(infoSection);
  }

  return builder;
}

/**
 * Generates a standardized premium footer section with action buttons for tool homepages.
 * Ensures configure/selector actions are always consistently placed at the bottom of the card.
 */
export function buildToolFooter(
  toolId: string,
  hasSettings: boolean
): GoogleAppsScript.Card_Service.CardSection {
  const section = CardService.newCardSection().setHeader('Navigation & Settings');

  if (hasSettings) {
    const settingsAction = CardService.newAction()
      .setFunctionName('buildUnifiedSettingsCard')
      .setParameters({ toolId });

    section.addWidget(
      CardService.newDecoratedText()
        .setStartIcon(CardService.newIconImage().setIconUrl(ICON_URLS.settings))
        .setText('<b>Configuration & Preferences</b>')
        .setBottomLabel('Customize how this tool operates')
        .setButton(
          CardService.newTextButton()
            .setText('Settings')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(settingsAction)
        )
    );
  }

  const backAction = CardService.newAction().setFunctionName('openToolSelector');
  section.addWidget(
    CardService.newDecoratedText()
      .setStartIcon(CardService.newIconImage().setIconUrl(ICON_URLS.info))
      .setText('<b>Sticky Assistant Home</b>')
      .setBottomLabel('Return to the main tool selector')
      .setButton(CardService.newTextButton().setText('Switch Tool').setOnClickAction(backAction))
  );

  return section;
}
