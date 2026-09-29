QUIZMASTER v4 - DYNAMIC EVIDENCE STORAGE
========================================

PURPOSE
-------
QuizMaster can now use one generic Apps Script gateway for many different
assessment collections. You choose the collection in submission-settings.json.
The gateway adds the current Auckland year automatically.

Example configuration:

  "storage": {
    "rootName": "PHS Building Evidence",
    "year": "auto"
  }

In 2026, the first real submission creates:

PHS Building Evidence - 2026
  Submission Register - 2026
  RNR - Mr Ranson
    123456
      123456_Student_Name_US24352_learner-self-reflection.pdf
      123456_Student_Name_US24352_assessment-record.pdf
      123456_Student_Name_US24352.puk

The unit standard is kept in the filenames and register rather than adding an
extra unit-standard folder.

CHANGING COLLECTIONS
--------------------
Only change storage.rootName in submission-settings.json.

Example:

  "rootName": "PHS Engineering Evidence"

The next real submission creates/uses:

  PHS Engineering Evidence - 2026

Changing it back to "PHS Building Evidence" reconnects to the existing
PHS Building Evidence - 2026 folder and register.

YEAR HANDLING
-------------
The year is determined server-side using Pacific/Auckland time.
Do not add a year to rootName. If one is accidentally added at the end, the
gateway removes it before creating the destination.

When the calendar year becomes 2027, the same configuration automatically uses:

  PHS Building Evidence - 2027

with a new:

  Submission Register - 2027

The 2026 evidence is left untouched.

REGISTER BEHAVIOUR
------------------
Each register belongs to ONE collection and ONE year.

A row is unique by:
- Teacher
- Unit Standard
- Student ID
- Assessment ID

Therefore Learner Self Reflection and Assessment Record are separate rows.
Resubmitting the same assessment updates only its own row and increments that
row's submission count.

The register-row cache is also scoped to the actual spreadsheet, so identical
student/assessment combinations in different evidence collections cannot clash.

SPEED BEHAVIOUR RETAINED
------------------------
This version keeps the v3 fast submission changes:
- PDF base64 prepared before Submit to Teacher is pressed
- PDF and .puk preparation run together
- fast cache-only confirmation checks
- expensive Drive save is outside the global lock
- only short duplicate-claim/register operations are locked
- folder IDs and register rows are cached
- no whole-column formatting on every submission

FIRST-TIME SETUP / UPGRADE
--------------------------
1. In Apps Script, replace Code.gs with:
     GoogleAppsScript-Code-QuizMaster-v4-Dynamic.gs

2. Save it.

3. Run setup() ONCE from the Apps Script editor and approve permissions.
   setup() deliberately creates NO evidence folder or register.

4. Deploy > Manage deployments > Edit.

5. Select New version and Deploy.

6. Keep:
     Execute as: Me
     Who has access: Anyone

7. Keep the existing /exec URL. The supplied submission-settings.json already
   contains the current QuizMaster gateway URL.

8. On GitHub replace:
     index.html
     script.js
     sw.js
     submission-settings.json

9. Do NOT replace questions.json.

10. Hard refresh QuizMaster once after GitHub Pages updates. The service-worker
    cache has been bumped to v19-dynamic to force the new front end to load.

USING ANOTHER SUBJECT / ASSESSMENT COLLECTION
---------------------------------------------
Edit only submission-settings.json, for example:

{
  "schemaVersion": 2,
  "gateway": {
    "url": "YOUR EXISTING /exec URL",
    "provider": "apps-script"
  },
  "storage": {
    "rootName": "PHS Food Technology Evidence",
    "year": "auto"
  },
  "features": {
    "submitPdf": true,
    "submitPuk": true
  }
}

The gateway currently requires rootName to begin with "PHS ". This is a small
server-side safeguard against arbitrary public folder creation. If you ever need
a non-PHS prefix, change ALLOWED_ROOT_PREFIX in Code.gs deliberately.

OLD FOLDERS
-----------
Existing folders such as PHS Building Submissions are not deleted or changed.
This version does not use the old global ROOT_FOLDER_ID or
REGISTER_SPREADSHEET_ID properties. New dynamic destinations use their own
collection+year property keys.

NOTIFICATIONS
-------------
Teacher email notification remains available in Code.gs but is OFF by default:

  NOTIFY_TEACHER: false

No Advanced Gmail service is required.
