/**
 * ============================================================================
 * Winton Wallet
 * Theme.gs
 * ----------------------------------------------------------------------------
 * Application theme, colors, fonts, and formatting helpers.
 * ============================================================================
 */

const WW_THEME = Object.freeze({

  COLORS: Object.freeze({

    PRIMARY: "#8E44AD",      // Purple
    PINK: "#FF5FA2",
    TEAL: "#00B8C9",
    LIME: "#9CD93A",
    ORANGE: "#FF9F1C",

    WHITE: "#FFFFFF",
    LIGHT_GRAY: "#F3F4F6",
    GRAY: "#D1D5DB",
    DARK_GRAY: "#4B5563",

    SUCCESS: "#2ECC71",
    WARNING: "#F4B400",
    ERROR: "#E74C3C",

    TEXT: "#2D3436"

  }),

  FONT: Object.freeze({

    FAMILY: "Arial",
    HEADER_SIZE: 14,
    SUBHEADER_SIZE: 12,
    BODY_SIZE: 10

  })

});

/**
 * Applies the standard title style.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 */
function wwFormatTitle(range) {

  range
    .setFontFamily(WW_THEME.FONT.FAMILY)
    .setFontSize(WW_THEME.FONT.HEADER_SIZE)
    .setFontWeight("bold")
    .setFontColor(WW_THEME.COLORS.WHITE)
    .setBackground(WW_THEME.COLORS.PRIMARY);

}

/**
 * Applies the standard header style.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 */
function wwFormatHeader(range) {

  range
    .setFontFamily(WW_THEME.FONT.FAMILY)
    .setFontSize(WW_THEME.FONT.SUBHEADER_SIZE)
    .setFontWeight("bold")
    .setFontColor(WW_THEME.COLORS.TEXT)
    .setBackground(WW_THEME.COLORS.GRAY);

}

/**
 * Applies the standard body style.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 */
function wwFormatBody(range) {

  range
    .setFontFamily(WW_THEME.FONT.FAMILY)
    .setFontSize(WW_THEME.FONT.BODY_SIZE)
    .setFontColor(WW_THEME.COLORS.TEXT)
    .setBackground(WW_THEME.COLORS.WHITE);

}