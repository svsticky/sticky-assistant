// Google Apps Script Trigger Entrypoints
export * from './triggers';

// Core UI & Selection Navigation Callbacks
export { APP_VERSION } from './core/version';
export { openTool, openToolSelector } from './core/triggerHandler';
export {
  buildUnifiedSettingsCard,
  saveIndividualSetting,
  deleteWorkspaceOverride,
} from './core/settingsCard';

// Action Points Extractor Callbacks
export {
  scanDocument,
  sendToTodoist,
  applyDocumentActions,
  populatePeopleConfig,
  jumpToTask,
  buildAddPeopleLayerCard,
} from './tools/action-points-extractor/cards';

// PDF Comment Viewer Callbacks
export { exportCommentsToSheet } from './tools/comments-extractor/cards';

// Committee Creator Callbacks
export { onDriveSelection, runScan, runExecution } from './tools/committee-creator/editing';

// Calendar Sync Callbacks & Triggers
export {
  openCreateJobCard,
  openEditJobCard,
  saveJobAction,
  triggerSyncManual,
  openDeleteJobConfirmCard,
  deleteJobAction,
} from './tools/calendar-sync/cards';
export {
  handleHourlySync,
  handleBackgroundInitialSync,
  handleBackgroundDeletion,
} from './tools/calendar-sync/sync';

// Properties Debugger Callbacks
export {
  openEditPropertyCard,
  savePropertyAction,
  openDeletePropertyConfirmCard,
  deletePropertyAction,
} from './tools/properties-debugger/cards';
