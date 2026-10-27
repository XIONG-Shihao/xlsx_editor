/** Foreground tests are deliberately separate from normal background checks. */
export function assertGuiAllowed(env = process.env) {
  if (env.PLATFORM_GUI_TESTS !== '1') {
    throw new Error(
      'Foreground GUI tests are disabled. Arrange a desktop testing session first, then set PLATFORM_GUI_TESTS=1 for that command only. Use --list to inspect tests without opening windows.'
    )
  }
}
