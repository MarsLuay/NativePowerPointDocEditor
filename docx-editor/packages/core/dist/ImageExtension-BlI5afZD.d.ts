import { WrapType } from './docx/wrapTypes.js';

/**
 * Anchored wrap-type targets — the OOXML wrap types except `inline`.
 */
type AnchorWrapType = Exclude<WrapType, 'inline'>;
/**
 * User-facing layout choices that mirror Word's Wrap Text menu. These are
 * directional shortcuts that fold (`wrapType`, `cssFloat`) into a single
 * picker option:
 *
 *   - `squareLeft`  — image floats left, text wraps on the right
 *   - `squareRight` — image floats right, text wraps on the left
 *   - `inline`      — image flows in the line as a glyph
 *   - the other targets are 1:1 with their OOXML wrap types
 *
 * `setImageWrapType` accepts both raw OOXML targets (square / tight / through
 * / topAndBottom / behind / inFront / inline) and the directional convenience
 * targets.
 */
type ImageLayoutTarget = AnchorWrapType | 'squareLeft' | 'squareRight' | 'inline';
/**
 * Optional context for `setImageWrapType` transitions:
 *
 *   - `initialPositionEmu`: when promoting an inline image to an anchor, the
 *     caller passes the inline image's current rendered position relative to
 *     the column origin in EMUs. The command stores this as the anchor's
 *     `wp:positionH` / `wp:positionV` so Word renders the new float exactly
 *     where the inline image used to sit (Word's own behavior).
 */
interface SetImageWrapTypeOptions {
    initialPositionEmu?: {
        horizontalEmu: number;
        verticalEmu: number;
    };
}

export type { AnchorWrapType as A, ImageLayoutTarget as I, SetImageWrapTypeOptions as S };
