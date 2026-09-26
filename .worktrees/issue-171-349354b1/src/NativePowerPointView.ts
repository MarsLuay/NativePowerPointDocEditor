/**
 * Compatibility entry point.
 *
 * Obsidian registrations and third-party callers historically import this
 * module. The implementation belongs to the PowerPoint UI boundary.
 */
export {
  NativePowerPointView,
  NATIVE_POWERPOINT_VIEW_TYPE,
  MODERN_POWERPOINT_EXTENSIONS,
  LEGACY_POWERPOINT_EXTENSIONS,
  MACRO_ENABLED_POWERPOINT_EXTENSIONS,
  EDITABLE_POWERPOINT_EXTENSIONS,
  POWERPOINT_EXTENSIONS,
  isPowerPointExtension,
  isModernPowerPointExtension,
  isEditablePowerPointExtension,
  isMacroEnabledPowerPointExtension,
} from './powerpoint/ui/NativePowerPointView';
