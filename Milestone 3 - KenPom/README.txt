===========================================================================
 MILESTONE 3  -  KENPOM AUTOMATION  (final step)
===========================================================================

WHAT THIS IS
  This is the same College Basketball tool from Milestone 2, now with the
  four KenPom columns filled in automatically as well:

      KenPom SOS, Overall Rating, Offense (ORtg), Defense (DRtg)

  Replace your Milestone 2 basketball folder with this one - everything
  else works exactly the same.

ONE-TIME SETUP  (your KenPom login)
  KenPom needs a paid subscription (about $25/year at kenpom.com) and a
  login - it has no public data feed, so the tool signs in as you to read
  the numbers.

  1. Open config.py in Notepad.
  2. Put your KenPom email and password on these two lines:
         KENPOM_EMAIL = "you@example.com"
         KENPOM_PASSWORD = "your-password"
  3. Save and close.

  That login stays on your own PC in this folder - it is only used to sign
  in to KenPom.

HOW TO UPDATE  (same as before)
  1. Close College Basketball.xlsx if it's open.
  2. Double-click "Update Basketball.bat".
  3. It now also fills the four KenPom columns. Done.

NOTES
  * The first run installs one extra helper for the KenPom login - that's
    automatic, just leave the window open the first time.
  * If you ever change your KenPom password, update it in config.py too.
  * KenPom occasionally changes its site; if a KenPom pull stops working
    one week, let me know and I'll sort it (first fix is on me).
  * Leave the KenPom login blank and the tool simply goes back to leaving
    those four columns for manual entry.
===========================================================================
