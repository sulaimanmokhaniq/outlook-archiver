"""Words the program writes itself: progress, problems, README, Excel list.

The screen's own words live in web/app.js. Both come in English and Arabic.
"""
from __future__ import annotations

LANGS = ("en", "ar")

MESSAGES: dict[str, dict[str, str]] = {
    "en": {
        # item kinds (used for folder names and the Excel list)
        "kind_mail": "Emails",
        "kind_meeting": "Meeting invitations",
        "kind_appointment": "Calendar",
        "kind_contact": "Contacts",
        "kind_task": "Tasks",
        "kind_note": "Notes",
        "kind_other": "Other",
        "no_date": "No date",
        "files_dir": "Files",
        "attachments_dir": "{name} - attachments",
        "no_subject": "(no subject)",
        "no_name": "(no name)",
        # Excel list
        "col_type": "Type",
        "col_folder": "Outlook folder",
        "col_date": "Date",
        "col_from": "From",
        "col_to": "To / CC",
        "col_subject": "Subject",
        "col_categories": "Categories",
        "col_attachments": "Attachments",
        "col_files": "Saved files",
        "col_removed": "Removed from mailbox",
        "yes": "Yes",
        "index_file": "Archive list.csv",
        "readme_file": "README.txt",
        # progress
        "p_connecting": "Connecting to Outlook ...",
        "p_reading_tree": "Reading the folder list ...",
        "p_looking": "Looking in {folder} ...",
        "p_looking_n": "Looking in {folder} ... {n:,} of {total:,}",
        "p_found": "Finished looking: {n:,} items match.",
        "p_pst": "Creating the PST file ...",
        "p_archiving": "Archiving {i:,} of {n:,}: {subject}",
        "p_checking": "Checking the PST file before removing anything ...",
        "p_removing": "Removing from the mailbox: {i:,} of {n:,}",
        "p_writing": "Writing the Excel list and README ...",
        "p_done": "Done.",
        "p_stopped": "Stopped.",
        # connection
        "connected": "Connected to Outlook - {n} mailbox(es) / data file(s).",
        "demo": "DEMO - made-up data, not your real Outlook.",
        # problems
        "e_open": "Could not open {what}: {err}",
        "e_pst_copy": "Could not copy {what} into the PST: {err}",
        "e_save": "Could not save {what} as {fmt}: {err}",
        "e_atts": "Could not save the attachments of {what}: {err}",
        "e_pst_close": "Could not close the PST file cleanly: {err}",
        "e_no_word": "PDF files were skipped: Microsoft Word is not installed on this laptop.",
        "e_pst_create": "Outlook could not create the PST file {path}: {err}",
        "e_pst_open": "Outlook created the PST file but could not open it.",
        "e_skip_folder": "Skipped {folder}: {err}",
        "e_remove": "Could not remove {what} from the mailbox: {err}",
        "e_verify": "Nothing was removed from the mailbox, because the PST check failed: {detail}",
        "v_missing": "the PST folder '{folder}' has {found:,} items but {expected:,} were copied",
        "e_not_connected": "Not connected to Outlook yet. Click 'Reconnect'.",
        "e_unexpected": "Unexpected problem: {err}",
        "e_busy": "Please wait - the program is still working.",
        "e_stale": "The list of items is out of date. Count the items again.",
        "e_no_outlook": (
            "This program could not reach classic Outlook on this laptop.\n\n"
            "It works with classic Outlook (the version with a 'File' menu at the top left).\n"
            "* If Outlook has a 'New Outlook' switch at the top right that is ON, turn it OFF, "
            "then click 'Reconnect'.\n"
            "* If classic Outlook is not installed, install the Microsoft 365 desktop apps "
            "(Outlook is included), open Outlook once to sign in, then click 'Reconnect'."
        ),
        # checks on the choices
        "v_no_dir": "Choose a folder to save the archive in.",
        "v_bad_dir": "This folder does not exist: {path}",
        "v_no_name": "Give the archive a name.",
        "v_no_output": "Choose at least one way to save (PST, files or the Excel list).",
        "v_move_needs_pst": "Removing items from the mailbox needs the PST option to be on.",
        # README
        "r_title": "Outlook archive: {name}",
        "r_made": "Made on {when}",
        "r_count": "Items archived: {done:,} of {total:,} chosen",
        "r_stopped": " (stopped before the end)",
        "r_inside": "What is in this folder",
        "r_pst": ("* {file}\n  An Outlook data file. To open it in Outlook:\n"
                  "  File > Open & Export > Open Outlook Data File, then choose this file.\n"
                  "  It works with no internet connection."),
        "r_files": "* {dir}\\\n  Each item saved as its own file, sorted into folders.",
        "r_index": "* {file}\n  A list of every archived item. Opens in Excel; you can sort and filter it.",
        "r_problems": "Problems ({n})",
        "r_copies": "Nothing was deleted from Outlook: these are copies.",
        "r_moved": ("{n:,} items were removed from the mailbox after they were saved in the PST. "
                    "They are in 'Deleted Items' until it is emptied."),
        "r_kept": ("{n:,} items were left in the mailbox on purpose: contacts, tasks, "
                   "recurring meetings and meetings that have not happened yet."),
    },
    "ar": {
        "kind_mail": "البريد",
        "kind_meeting": "دعوات الاجتماعات",
        "kind_appointment": "التقويم",
        "kind_contact": "جهات الاتصال",
        "kind_task": "المهام",
        "kind_note": "الملاحظات",
        "kind_other": "أخرى",
        "no_date": "بدون تاريخ",
        "files_dir": "الملفات",
        "attachments_dir": "{name} - المرفقات",
        "no_subject": "(بدون موضوع)",
        "no_name": "(بدون اسم)",
        "col_type": "النوع",
        "col_folder": "مجلد Outlook",
        "col_date": "التاريخ",
        "col_from": "من",
        "col_to": "إلى / نسخة",
        "col_subject": "الموضوع",
        "col_categories": "التصنيفات",
        "col_attachments": "المرفقات",
        "col_files": "الملفات المحفوظة",
        "col_removed": "حُذف من صندوق البريد",
        "yes": "نعم",
        "index_file": "قائمة الأرشيف.csv",
        "readme_file": "اقرأني.txt",
        "p_connecting": "جارٍ الاتصال بـ Outlook ...",
        "p_reading_tree": "جارٍ قراءة قائمة المجلدات ...",
        "p_looking": "جارٍ البحث في {folder} ...",
        "p_looking_n": "جارٍ البحث في {folder} ... {n:,} من {total:,}",
        "p_found": "انتهى البحث: {n:,} عنصرًا مطابقًا.",
        "p_pst": "جارٍ إنشاء ملف PST ...",
        "p_archiving": "جارٍ أرشفة {i:,} من {n:,}: {subject}",
        "p_checking": "جارٍ فحص ملف PST قبل حذف أي شيء ...",
        "p_removing": "جارٍ الحذف من صندوق البريد: {i:,} من {n:,}",
        "p_writing": "جارٍ كتابة قائمة Excel وملف اقرأني ...",
        "p_done": "تم.",
        "p_stopped": "تم الإيقاف.",
        "connected": "متصل بـ Outlook - {n} صندوق بريد / ملف بيانات.",
        "demo": "تجربة - بيانات وهمية وليست بريدك الحقيقي.",
        "e_open": "تعذّر فتح {what}: {err}",
        "e_pst_copy": "تعذّر نسخ {what} إلى ملف PST: {err}",
        "e_save": "تعذّر حفظ {what} بصيغة {fmt}: {err}",
        "e_atts": "تعذّر حفظ مرفقات {what}: {err}",
        "e_pst_close": "تعذّر إغلاق ملف PST بشكل سليم: {err}",
        "e_no_word": "لم تُحفظ ملفات PDF لأن Microsoft Word غير مثبت على هذا الجهاز.",
        "e_pst_create": "تعذّر على Outlook إنشاء ملف PST ‏{path}: {err}",
        "e_pst_open": "أنشأ Outlook ملف PST لكنه لم يتمكن من فتحه.",
        "e_skip_folder": "تم تخطي {folder}: {err}",
        "e_remove": "تعذّر حذف {what} من صندوق البريد: {err}",
        "e_verify": "لم يُحذف أي شيء من صندوق البريد لأن فحص ملف PST لم ينجح: {detail}",
        "v_missing": "المجلد '{folder}' في ملف PST فيه {found:,} عنصرًا بينما نُسخ إليه {expected:,}",
        "e_not_connected": "لم يتم الاتصال بـ Outlook بعد. اضغط 'إعادة الاتصال'.",
        "e_unexpected": "حدثت مشكلة غير متوقعة: {err}",
        "e_busy": "الرجاء الانتظار - البرنامج ما زال يعمل.",
        "e_stale": "قائمة العناصر لم تعد محدّثة. اضغط 'عدّ العناصر' مرة أخرى.",
        "e_no_outlook": (
            "لم يتمكن البرنامج من الوصول إلى Outlook الكلاسيكي على هذا الجهاز.\n\n"
            "البرنامج يعمل مع Outlook الكلاسيكي (النسخة التي فيها قائمة 'ملف' أعلى اليسار).\n"
            "* إذا كان في أعلى Outlook زر 'Outlook الجديد' وهو مفعّل، أوقفه ثم اضغط 'إعادة الاتصال'.\n"
            "* إذا لم يكن Outlook الكلاسيكي مثبتًا، ثبّت تطبيقات Microsoft 365 لسطح المكتب "
            "(ومنها Outlook)، وافتح Outlook مرة لتسجيل الدخول، ثم اضغط 'إعادة الاتصال'."
        ),
        "v_no_dir": "اختر مجلدًا لحفظ الأرشيف فيه.",
        "v_bad_dir": "هذا المجلد غير موجود: {path}",
        "v_no_name": "اكتب اسمًا للأرشيف.",
        "v_no_output": "اختر طريقة حفظ واحدة على الأقل (PST أو ملفات أو قائمة Excel).",
        "v_move_needs_pst": "الحذف من صندوق البريد يتطلب تفعيل خيار ملف PST.",
        "r_title": "أرشيف Outlook: {name}",
        "r_made": "تاريخ الإنشاء: {when}",
        "r_count": "العناصر المؤرشفة: {done:,} من {total:,}",
        "r_stopped": " (تم الإيقاف قبل النهاية)",
        "r_inside": "محتويات هذا المجلد",
        "r_pst": ("* {file}\n  ملف بيانات Outlook. لفتحه في Outlook:\n"
                  "  ملف > فتح وتصدير > فتح ملف بيانات Outlook، ثم اختر هذا الملف.\n"
                  "  يعمل بدون اتصال بالإنترنت."),
        "r_files": "* {dir}\\\n  كل عنصر محفوظ في ملف مستقل، ومرتّب في مجلدات.",
        "r_index": "* {file}\n  قائمة بكل العناصر المؤرشفة. تُفتح في Excel ويمكن ترتيبها وتصفيتها.",
        "r_problems": "مشكلات ({n})",
        "r_copies": "لم يُحذف أي شيء من Outlook: هذه نسخ.",
        "r_moved": ("حُذف {n:,} عنصرًا من صندوق البريد بعد حفظها في ملف PST. "
                    "وهي موجودة في 'العناصر المحذوفة' إلى أن يتم إفراغها."),
        "r_kept": ("بقي {n:,} عنصرًا في صندوق البريد عمدًا: جهات الاتصال والمهام "
                   "والاجتماعات المتكررة والاجتماعات التي لم يحن موعدها."),
    },
}


class Translator:
    def __init__(self, lang: str = "en") -> None:
        self.lang = lang if lang in MESSAGES else "en"

    def __call__(self, key: str, **kw) -> str:
        text = MESSAGES[self.lang].get(key) or MESSAGES["en"][key]
        return text.format(**kw) if kw else text
