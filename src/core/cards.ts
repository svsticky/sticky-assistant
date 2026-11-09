import { COLORS, getAppName } from './branding';

/**
 * Builds a standardized, beautiful status or message card.
 * Normalizes user feedback and removes redundant message card code.
 */
export function buildStatusCard(
  title: string,
  message: string,
  statusType: 'info' | 'success' | 'warning' | 'error' = 'info'
): GoogleAppsScript.Card_Service.Card {
  const builder = CardService.newCardBuilder();
  builder.setHeader(CardService.newCardHeader().setTitle(getAppName()));

  let icon: GoogleAppsScript.Card_Service.Icon;
  let color: string;

  switch (statusType) {
    case 'success':
      icon = CardService.Icon.STAR;
      color = COLORS.SUCCESS;
      break;
    case 'error':
      icon = CardService.Icon.OFFER;
      color = COLORS.ERROR;
      break;
    case 'warning':
      icon = CardService.Icon.CLOCK;
      color = COLORS.WARNING;
      break;
    case 'info':
    default:
      icon = CardService.Icon.DESCRIPTION;
      color = COLORS.PRIMARY;
      break;
  }

  const section = CardService.newCardSection().addWidget(
    CardService.newDecoratedText()
      .setStartIcon(CardService.newIconImage().setIcon(icon))
      .setText(`<font color="${color}"><b>${title}</b></font>`)
      .setBottomLabel(message)
      .setWrapText(true)
  );

  const backAction = CardService.newAction().setFunctionName('openToolSelector');
  section.addWidget(
    CardService.newTextButton()
      .setText('Back to Home')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(backAction)
  );

  builder.addSection(section);
  return builder.build();
}
