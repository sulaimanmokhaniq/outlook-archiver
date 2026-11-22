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
        "col_check": "Check",
        "chk_ok": "OK",
        "chk_not_copied": "Not copied",
        "chk_not_in_pst": "Not found in the PST",
        "chk_file_missing": "File missing",
        "chk_stopped": "Not done (stopped)",
        "chk_not_checked": "Could not be checked",
        "p_verifying": "Checking the archive: every item and its size ...",
        "e_check": "Could not check the PST: {err}",
        "v_missing_n": "{n:,} item(s) could not be found in the PST",
        "r_check": "Check",
        "r_chk_items": "Items chosen: {n:,} ({size})",
        "r_chk_pst": "Found in the PST: {found:,} of {n:,} (size in Outlook {size_a}, in the PST {size_b})",
        "r_chk_files": "Files saved and readable: {found:,} of {n:,}",
        "r_chk_ok": "Result: every chosen item is in the archive.",
        "r_chk_bad": "Result: {n:,} item(s) are missing or incomplete - see below and the '{col}' column in the Excel list.",
        "r_missing": "Missing or incomplete",
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
        "v_move_needs_pst": ("Removing items from the mailbox needs a safe copy first: the PST file, "
                             "or (in Microsoft 365 online mode) the original files."),
        "connected_online": "Signed in to Microsoft 365 as {who}.",
        "p_signing_in": "Signing in to Microsoft 365 - finish the sign-in in the browser window that opened ...",
        "e_try_online": ("You can also use the Microsoft 365 online mode instead: it works with the new Outlook "
                         "too. Click 'Use Microsoft 365 online'."),
        "e_no_client_id": ("Microsoft 365 online mode needs an Application (client) ID before you can sign in. "
                           "Open the connection settings and paste the ID your IT team (or the README steps) gave you."),
        "e_signin": "Signing in to Microsoft 365 did not work: {err}",
        "e_network": "Could not reach Microsoft 365. Check the internet connection. ({err})",
        "e_denied": ("Microsoft 365 refused access: {err}\n\nYour IT team may need to approve this program "
                     "for your account (the README explains what to ask)."),
        "e_graph": "Microsoft 365 answered with an error ({code}): {err}",
        "e_no_pst_online": "PST files can only be made with classic Outlook. In Microsoft 365 online mode, save the original files instead.",
        "e_format_online": "{fmt} files cannot be made in Microsoft 365 online mode.",
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
        "col_check": "الفحص",
        "chk_ok": "سليم",
        "chk_not_copied": "لم يُنسخ",
        "chk_not_in_pst": "غير موجود في ملف PST",
        "chk_file_missing": "ملف ناقص",
        "chk_stopped": "لم يتم (تم الإيقاف)",
        "chk_not_checked": "تعذّر فحصه",
        "p_verifying": "جارٍ فحص الأرشيف: كل عنصر وحجمه ...",
        "e_check": "تعذّر فحص ملف PST: {err}",
        "v_missing_n": "تعذّر العثور على {n:,} عنصرًا في ملف PST",
        "r_check": "الفحص",
        "r_chk_items": "العناصر المختارة: {n:,} ({size})",
        "r_chk_pst": "الموجود في ملف PST: {found:,} من {n:,} (الحجم في Outlook {size_a}، وفي PST {size_b})",
        "r_chk_files": "الملفات المحفوظة والسليمة: {found:,} من {n:,}",
        "r_chk_ok": "النتيجة: كل العناصر المختارة موجودة في الأرشيف.",
        "r_chk_bad": "النتيجة: {n:,} عنصرًا ناقصًا أو غير مكتمل - انظر القائمة أدناه وعمود '{col}' في قائمة Excel.",
        "r_missing": "عناصر ناقصة أو غير مكتملة",
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
        "v_move_needs_pst": ("الحذف من صندوق البريد يتطلب نسخة آمنة أولًا: ملف PST، "
                             "أو (في وضع Microsoft 365 عبر الإنترنت) الملفات الأصلية."),
        "connected_online": "تم تسجيل الدخول إلى Microsoft 365 باسم {who}.",
        "p_signing_in": "جارٍ تسجيل الدخول إلى Microsoft 365 - أكمل تسجيل الدخول في نافذة المتصفح التي فُتحت ...",
        "e_try_online": ("يمكنك أيضًا استخدام وضع Microsoft 365 عبر الإنترنت بدلًا منه: يعمل مع Outlook الجديد "
                         "أيضًا. اضغط 'استخدم Microsoft 365 عبر الإنترنت'."),
        "e_no_client_id": ("وضع Microsoft 365 عبر الإنترنت يحتاج إلى معرّف التطبيق (Client ID) قبل تسجيل الدخول. "
                           "افتح إعدادات الاتصال والصق المعرّف الذي أعطاك إياه فريق تقنية المعلومات (أو خطوات ملف README)."),
        "e_signin": "لم ينجح تسجيل الدخول إلى Microsoft 365: {err}",
        "e_network": "تعذّر الوصول إلى Microsoft 365. تحقّق من الاتصال بالإنترنت. ({err})",
        "e_denied": ("رفض Microsoft 365 الوصول: {err}\n\nقد يحتاج فريق تقنية المعلومات إلى الموافقة على هذا البرنامج "
                     "لحسابك (ملف README يشرح ما يجب طلبه)."),
        "e_graph": "ردّ Microsoft 365 بخطأ ({code}): {err}",
        "e_no_pst_online": "ملفات PST تُنشأ فقط عبر Outlook الكلاسيكي. في وضع Microsoft 365 عبر الإنترنت، احفظ الملفات الأصلية بدلًا منها.",
        "e_format_online": "لا يمكن إنشاء ملفات {fmt} في وضع Microsoft 365 عبر الإنترنت.",
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
