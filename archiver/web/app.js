"use strict";
/* Outlook Archiver - the screen. Talks to the program running on this laptop. */

const TOKEN = new URLSearchParams(location.search).get("t") || "";
const $ = (id) => document.getElementById(id);

// --------------------------------------------------------------- words
const STR = {
  en: {
    app_title: "Outlook Archiver",
    app_sub: "Keep chosen emails, meetings and more on this laptop",
    lang_other: "العربية",
    conn_wait: "Connecting to Outlook…",
    conn_ok: "Outlook connected · {n} mailbox(es)",
    conn_demo: "Demo · made-up mailbox",
    conn_bad: "Outlook not connected",
    conn_online: "Microsoft 365 online · {who}",
    conn_bg: "Classic Outlook in the background · {n} mailbox(es)",
    bg_note: "You are using the new Outlook - no need to switch. This program runs classic Outlook hidden in the background to read the mailbox and make PST files, and closes it again when you close the program.",
    conn_title: "Connection",
    conn_intro: "Choose how the program reaches the mailbox. Your choice is remembered.",
    eng_classic: "Classic Outlook on this laptop",
    eng_classic_d: "Makes PST files. Works even if you use the new Outlook: classic Outlook runs hidden in the background, so you do not need to switch. Needs classic Outlook installed (it comes with the Microsoft 365 desktop apps).",
    eng_online: "Microsoft 365 online - works with the new Outlook",
    eng_online_d: "For laptops without classic Outlook. Reads the mailbox straight from Microsoft 365 after you sign in with your work account. Saves the original emails (.eml), calendar (.ics) and contacts (.vcf). It cannot make PST files.",
    eng_in_use: "in use",
    eng_unavailable: "not available on this computer",
    use_classic: "Use classic Outlook",
    use_online: "Use Microsoft 365 online",
    sign_out: "Sign out",
    client_id: "Application (client) ID",
    client_id_help: "A one-time setting from your IT team, or from an app registration in Microsoft Entra (the README has the steps). It is not a password.",
    tenant: "Organization (tenant) - leave as it is unless IT says otherwise",
    advanced: "Advanced",
    need_client_id: "Paste the Application (client) ID first.",
    fmt_eml: "Original files",
    reconnect: "Reconnect",
    s1: "Folders",
    s1_hint: "Tick where to look. Sub-folders follow.",
    search_folders: "Search folders…",
    tick_all: "Tick all",
    untick_all: "Untick all",
    tree_wait: "Reading your folders…",
    tree_noconn: "Outlook is not connected, so no folders can be shown.",
    tree_none: "No folders found.",
    tree_nomatch: "No folder matches your search.",
    s2: "What to archive",
    kind_mail: "Emails & drafts",
    kind_meeting: "Meeting invitations",
    kind_appointment: "Calendar",
    kind_contact: "Contacts",
    kind_task: "Tasks",
    kind_note: "Notes",
    kind_other: "Other",
    s3: "Time period",
    p_all: "All time",
    p_this_year: "This year",
    p_last_year: "Last year",
    p_last12: "Last 12 months",
    p_older1: "Older than 1 year",
    p_older2: "Older than 2 years",
    p_older3: "Older than 3 years",
    p_older5: "Older than 5 years",
    p_year: "A specific year",
    p_custom: "Custom dates",
    range_all: "Everything, whatever the date",
    range_both: "From {a} to {b}",
    range_until: "Everything up to {b}",
    range_since: "Everything from {a}",
    from: "From",
    to: "To",
    date_field: "Date to use:",
    df_auto: "Best date for each item (received / meeting date)",
    df_received: "Received date",
    df_sent: "Sent date",
    df_created: "Created date",
    df_modified: "Last changed date",
    s4: "Filters",
    optional: "optional",
    s4_active: "{n} on",
    f_sender: "From (sender) contains",
    f_recipients: "To / CC contains",
    f_subject: "Subject contains",
    f_category: "Category contains",
    f_body: "Words in the text",
    f_body_hint: "Slower: every item has to be opened.",
    f_attachments: "Attachments",
    att_any: "Any",
    att_with: "With attachments",
    att_without: "Without attachments",
    f_hint: "Separate several names or words with ;  - any one of them is enough.",
    s5: "Save as",
    o_pst: "PST file",
    o_pst_d: "Opens inside Outlook with the same folders. Works with no internet.",
    o_pst_show: "Show it in Outlook when done",
    o_files: "Also save each item as its own file",
    fmt_msg: "Outlook",
    fmt_mht: "Web page",
    fmt_pdf: "PDF",
    fmt_txt: "Text",
    fmt_pdf_note: "needs Word",
    o_atts: "Save attachments as normal files",
    o_index: "Excel list",
    o_index_d: "One row per item: date, sender, subject, folder. Easy to sort and search.",
    o_layout: "Sort files into",
    lay_folder_year: "Outlook folder › year",
    lay_type_year: "Item type › year",
    lay_year_month: "Year › month",
    lay_flat: "One folder",
    o_name: "Archive name",
    o_dir: "Save in this folder",
    browse: "Browse…",
    name_base: "Outlook archive",
    name_until: "until {d}",
    name_since: "from {d}",
    s6: "Repeats & clean-up",
    o_skip: "Skip items that were already archived",
    hist_some: "{n} items are in earlier archives on this laptop.",
    hist_none: "Nothing archived yet.",
    forget: "Forget",
    forget_q: "Forget which items were archived before?\nThe archives themselves are not touched.",
    o_move: "Remove items from the mailbox after archiving",
    o_move_d: "Frees up mailbox space. Nothing is removed until the PST file has been checked, and removed items go to Deleted Items, so they can be restored. Contacts, tasks, recurring and upcoming meetings always stay.",
    o_move_needs_pst: "Turn on the PST file option first.",
    o_move_needs_eml: "Turn on “Original files” in step 5 first - that is the safe copy in online mode.",
    sum_none: "Tick at least one folder to start.",
    sum_folders: "{n} folder(s)",
    sum_all_types: "all item types",
    sum_types: "{n} item types",
    sum_filters: "{n} filter(s)",
    count_btn: "Count items",
    archive_btn: "Start archiving",
    count_none: "Click “Count items” to see how many items match, before anything is copied.",
    count_res: "{n} items match",
    count_one: "1 item matches",
    count_size: "{s} in total",
    col_size: "Size",
    pv_total: "Total size: {s}",
    space_title: "Space - now and after archiving",
    space_calc: "Calculating…",
    sp_now: "Now",
    sp_after: "After",
    sp_items: "Chosen items",
    sp_items_v: "{s} in Outlook",
    sp_disk: "Free space on this laptop",
    sp_disk_d: "Free space on drive {d}",
    sp_archive: "The archive will take about {s} on the laptop.",
    sp_mailbox: "Mailbox “{n}”",
    sp_nochange: "no change - copies only",
    sp_after_empty: "once Deleted Items is emptied",
    sp_file: "Outlook’s file on this laptop ({ext})",
    sp_file_shrink: "gets smaller after “Compact Now”",
    sp_unknown: "unknown",
    sp_not_enough: "Not enough free space: the archive needs about {need}, but only {free} is free. Choose another folder or drive, or save fewer kinds of files.",
    approx: "about {s}",
    done_size: "archive size",
    done_title_bad: "Finished - but some items are missing",
    chk_title: "Check: is everything in the archive?",
    chk_items: "Items: {a} chosen · {b} confirmed in the archive",
    chk_pst: "In the PST: {a} of {b} found · size {s1} in Outlook, {s2} in the PST",
    chk_files: "Files: {a} of {b} saved and readable",
    chk_all_ok: "Every chosen item is in the archive - nothing is missing.",
    chk_bad: "{n} item(s) are missing or incomplete. They are listed below, in the README and in the Excel list (column “Check”).",
    chk_show: "Show the missing items ({n})",
    chk_size_note: "Small size differences are normal: a PST stores items in its own way.",
    sp_items_n: "{n} items",
    done_freed: "freed from mailbox",
    conf_one: "1 item",
    count_skipped: "{n} already archived, skipped",
    count_stale: "You changed your choices - count again.",
    show_list: "Show the list",
    need_folders: "Tick at least one folder in step 1.",
    need_kinds: "Choose at least one kind of item in step 2.",
    need_dates: "The ‘From’ date is after the ‘To’ date.",
    need_output: "Choose at least one way to save in step 5.",
    need_dir: "Choose a folder to save the archive in.",
    nothing_match: "No items match these choices, so there is nothing to archive.",
    preview_title: "Matching items",
    by_type: "By type",
    by_year: "By year",
    col_date: "Date",
    col_type: "Type",
    col_from: "From",
    col_subject: "Subject",
    col_folder: "Folder",
    pv_more: "Showing the first {shown} of {total}. The Excel list will have them all.",
    no_subject: "(no subject)",
    confirm_title: "Ready to archive",
    conf_items: "{n} items",
    conf_folders: "Folders",
    conf_period: "Period",
    conf_types: "Types",
    conf_as: "Saved as",
    conf_where: "Saved in",
    conf_skipped: "Already archived",
    conf_skipped_v: "{n} items skipped",
    what_pst: "PST file",
    what_files: "files ({list})",
    what_atts: "attachments",
    what_index: "Excel list",
    confirm_move: "After the copy is checked, these items will be removed from the mailbox (moved to Deleted Items). Contacts, tasks, recurring and upcoming meetings stay.",
    confirm_move_ok: "I understand that items will be removed from the mailbox",
    confirm_copy: "Nothing will be deleted from Outlook. Keep Outlook open while this runs.",
    cancel: "Cancel",
    start: "Start",
    ok: "OK",
    close: "Close",
    job_connect: "Connecting to Outlook",
    job_count: "Looking for matching items",
    job_archive: "Archiving",
    elapsed: "Time so far: {t}",
    remaining: "About {t} left",
    stop: "Stop",
    stopping: "Stopping…",
    stopped: "Stopped. Nothing more was done.",
    done_title: "Your archive is ready",
    done_stopped: "Stopped part way",
    done_items: "items archived",
    done_files: "files saved",
    done_removed: "removed from mailbox",
    done_kept: "kept in mailbox",
    done_problems: "problems",
    done_where: "Saved in:",
    problems_title: "Show the {n} problem(s)",
    open_folder: "Open archive folder",
    sec: "{n} s",
    min: "{m} min {s} s",
    hour: "{h} h {m} min",
  },
  ar: {
    app_title: "أرشفة Outlook",
    app_sub: "احفظ ما تختاره من الرسائل والاجتماعات وغيرها على هذا الجهاز",
    lang_other: "English",
    conn_wait: "جارٍ الاتصال بـ Outlook…",
    conn_ok: "Outlook متصل · {n} صندوق بريد",
    conn_demo: "تجربة · صندوق بريد وهمي",
    conn_bad: "Outlook غير متصل",
    conn_online: "Microsoft 365 عبر الإنترنت · {who}",
    conn_bg: "Outlook الكلاسيكي في الخلفية · {n} صندوق بريد",
    bg_note: "أنت تستخدم Outlook الجديد - لا حاجة للتبديل. يشغّل البرنامج Outlook الكلاسيكي مخفيًا في الخلفية لقراءة البريد وإنشاء ملفات PST، ويغلقه عند إغلاق البرنامج.",
    conn_title: "الاتصال",
    conn_intro: "اختر طريقة وصول البرنامج إلى صندوق البريد. سيتم تذكّر اختيارك.",
    eng_classic: "Outlook الكلاسيكي على هذا الجهاز",
    eng_classic_d: "ينشئ ملفات PST. يعمل حتى لو كنت تستخدم Outlook الجديد: يعمل Outlook الكلاسيكي مخفيًا في الخلفية فلا تحتاج إلى التبديل. يتطلب أن يكون Outlook الكلاسيكي مثبتًا (يأتي مع تطبيقات Microsoft 365 لسطح المكتب).",
    eng_online: "Microsoft 365 عبر الإنترنت - يعمل مع Outlook الجديد",
    eng_online_d: "للأجهزة التي ليس عليها Outlook الكلاسيكي. يقرأ صندوق البريد مباشرة من Microsoft 365 بعد تسجيل الدخول بحساب العمل. يحفظ الرسائل الأصلية (.eml) والتقويم (.ics) وجهات الاتصال (.vcf). لا يمكنه إنشاء ملفات PST.",
    eng_in_use: "مستخدم الآن",
    eng_unavailable: "غير متاح على هذا الجهاز",
    use_classic: "استخدم Outlook الكلاسيكي",
    use_online: "استخدم Microsoft 365 عبر الإنترنت",
    sign_out: "تسجيل الخروج",
    client_id: "معرّف التطبيق (Client ID)",
    client_id_help: "إعداد لمرة واحدة من فريق تقنية المعلومات، أو من تسجيل تطبيق في Microsoft Entra (الخطوات في ملف README). ليس كلمة مرور.",
    tenant: "المؤسسة (Tenant) - اتركها كما هي ما لم يطلب فريق تقنية المعلومات غير ذلك",
    advanced: "خيارات متقدمة",
    need_client_id: "الصق معرّف التطبيق (Client ID) أولًا.",
    fmt_eml: "الملفات الأصلية",
    reconnect: "إعادة الاتصال",
    s1: "المجلدات",
    s1_hint: "حدّد أين يتم البحث. المجلدات الفرعية تتبع المجلد الرئيسي.",
    search_folders: "ابحث عن مجلد…",
    tick_all: "تحديد الكل",
    untick_all: "إلغاء الكل",
    tree_wait: "جارٍ قراءة مجلداتك…",
    tree_noconn: "Outlook غير متصل، لذلك لا يمكن عرض المجلدات.",
    tree_none: "لم يتم العثور على مجلدات.",
    tree_nomatch: "لا يوجد مجلد بهذا الاسم.",
    s2: "ماذا تريد أن تؤرشف",
    kind_mail: "الرسائل والمسودات",
    kind_meeting: "دعوات الاجتماعات",
    kind_appointment: "التقويم",
    kind_contact: "جهات الاتصال",
    kind_task: "المهام",
    kind_note: "الملاحظات",
    kind_other: "أخرى",
    s3: "الفترة الزمنية",
    p_all: "كل الفترات",
    p_this_year: "هذه السنة",
    p_last_year: "السنة الماضية",
    p_last12: "آخر 12 شهرًا",
    p_older1: "أقدم من سنة",
    p_older2: "أقدم من سنتين",
    p_older3: "أقدم من 3 سنوات",
    p_older5: "أقدم من 5 سنوات",
    p_year: "سنة محددة",
    p_custom: "تواريخ مخصّصة",
    range_all: "كل العناصر مهما كان تاريخها",
    range_both: "من {a} إلى {b}",
    range_until: "كل ما قبل {b} وحتى نهايته",
    range_since: "كل ما بعد {a} بدءًا منه",
    from: "من",
    to: "إلى",
    date_field: "التاريخ المعتمد:",
    df_auto: "أنسب تاريخ لكل عنصر (الاستلام / موعد الاجتماع)",
    df_received: "تاريخ الاستلام",
    df_sent: "تاريخ الإرسال",
    df_created: "تاريخ الإنشاء",
    df_modified: "تاريخ آخر تعديل",
    s4: "عوامل التصفية",
    optional: "اختياري",
    s4_active: "{n} مفعّل",
    f_sender: "المرسل يحتوي",
    f_recipients: "إلى / نسخة يحتوي",
    f_subject: "الموضوع يحتوي",
    f_category: "التصنيف يحتوي",
    f_body: "كلمات في نص الرسالة",
    f_body_hint: "أبطأ: يجب فتح كل عنصر.",
    f_attachments: "المرفقات",
    att_any: "الكل",
    att_with: "مع مرفقات",
    att_without: "بدون مرفقات",
    f_hint: "افصل بين الأسماء أو الكلمات بـ ;  - ويكفي أن يطابق أحدها.",
    s5: "طريقة الحفظ",
    o_pst: "ملف PST",
    o_pst_d: "يُفتح داخل Outlook بنفس المجلدات، ويعمل بدون إنترنت.",
    o_pst_show: "أظهره في Outlook عند الانتهاء",
    o_files: "احفظ أيضًا كل عنصر في ملف مستقل",
    fmt_msg: "Outlook",
    fmt_mht: "صفحة ويب",
    fmt_pdf: "PDF",
    fmt_txt: "نص",
    fmt_pdf_note: "يتطلب Word",
    o_atts: "احفظ المرفقات كملفات عادية",
    o_index: "قائمة Excel",
    o_index_d: "صف لكل عنصر: التاريخ والمرسل والموضوع والمجلد. سهلة الترتيب والبحث.",
    o_layout: "ترتيب الملفات حسب",
    lay_folder_year: "مجلد Outlook › السنة",
    lay_type_year: "نوع العنصر › السنة",
    lay_year_month: "السنة › الشهر",
    lay_flat: "مجلد واحد",
    o_name: "اسم الأرشيف",
    o_dir: "احفظ في هذا المجلد",
    browse: "استعراض…",
    name_base: "أرشيف Outlook",
    name_until: "حتى {d}",
    name_since: "من {d}",
    s6: "التكرار والتنظيف",
    o_skip: "تخطَّ العناصر التي سبقت أرشفتها",
    hist_some: "{n} عنصرًا موجودة في أرشيفات سابقة على هذا الجهاز.",
    hist_none: "لم تتم أرشفة أي شيء بعد.",
    forget: "انسَ السجل",
    forget_q: "هل تريد نسيان العناصر التي تمت أرشفتها سابقًا؟\nلن تتأثر الأرشيفات نفسها.",
    o_move: "احذف العناصر من صندوق البريد بعد أرشفتها",
    o_move_d: "يوفّر مساحة في صندوق البريد. لا يُحذف شيء قبل فحص ملف PST، والعناصر المحذوفة تذهب إلى «العناصر المحذوفة» ويمكن استرجاعها. جهات الاتصال والمهام والاجتماعات المتكررة والقادمة تبقى دائمًا.",
    o_move_needs_pst: "فعّل خيار ملف PST أولًا.",
    o_move_needs_eml: "فعّل «الملفات الأصلية» في الخطوة 5 أولًا - فهي النسخة الآمنة في وضع الإنترنت.",
    sum_none: "حدّد مجلدًا واحدًا على الأقل للبدء.",
    sum_folders: "{n} مجلد",
    sum_all_types: "كل أنواع العناصر",
    sum_types: "{n} أنواع",
    sum_filters: "{n} تصفية",
    count_btn: "عدّ العناصر",
    archive_btn: "ابدأ الأرشفة",
    count_none: "اضغط «عدّ العناصر» لترى عدد العناصر المطابقة قبل نسخ أي شيء.",
    count_res: "{n} عنصرًا مطابقًا",
    count_one: "عنصر واحد مطابق",
    count_size: "الحجم الكلي {s}",
    col_size: "الحجم",
    pv_total: "الحجم الكلي: {s}",
    space_title: "المساحة - الآن وبعد الأرشفة",
    space_calc: "جارٍ الحساب…",
    sp_now: "الآن",
    sp_after: "بعد",
    sp_items: "العناصر المختارة",
    sp_items_v: "{s} في Outlook",
    sp_disk: "المساحة الفارغة على هذا الجهاز",
    sp_disk_d: "المساحة الفارغة على القرص {d}",
    sp_archive: "سيأخذ الأرشيف حوالي {s} على الجهاز.",
    sp_mailbox: "صندوق البريد «{n}»",
    sp_nochange: "بدون تغيير - نسخ فقط",
    sp_after_empty: "بعد إفراغ «العناصر المحذوفة»",
    sp_file: "ملف Outlook على هذا الجهاز ({ext})",
    sp_file_shrink: "يصغر بعد «ضغط الآن»",
    sp_unknown: "غير معروف",
    sp_not_enough: "لا توجد مساحة كافية: يحتاج الأرشيف حوالي {need}، والمتاح {free} فقط. اختر مجلدًا أو قرصًا آخر، أو قلّل أنواع الملفات المحفوظة.",
    approx: "حوالي {s}",
    done_size: "حجم الأرشيف",
    done_title_bad: "انتهى - لكن بعض العناصر ناقصة",
    chk_title: "الفحص: هل كل شيء موجود في الأرشيف؟",
    chk_items: "العناصر: {a} مختارة · {b} مؤكدة في الأرشيف",
    chk_pst: "في ملف PST: تم العثور على {a} من {b} · الحجم {s1} في Outlook و{s2} في PST",
    chk_files: "الملفات: {a} من {b} محفوظة وسليمة",
    chk_all_ok: "كل العناصر المختارة موجودة في الأرشيف - لا يوجد أي نقص.",
    chk_bad: "{n} عنصرًا ناقصًا أو غير مكتمل. تجدها أدناه وفي ملف اقرأني وفي قائمة Excel (عمود «الفحص»).",
    chk_show: "اعرض العناصر الناقصة ({n})",
    chk_size_note: "الفروق الصغيرة في الحجم طبيعية: ملف PST يخزّن العناصر بطريقته.",
    sp_items_n: "{n} عنصرًا",
    done_freed: "تم توفيره من صندوق البريد",
    conf_one: "عنصر واحد",
    count_skipped: "{n} سبقت أرشفتها وتم تخطيها",
    count_stale: "غيّرت اختياراتك - اضغط «عدّ العناصر» مرة أخرى.",
    show_list: "اعرض القائمة",
    need_folders: "حدّد مجلدًا واحدًا على الأقل في الخطوة 1.",
    need_kinds: "اختر نوعًا واحدًا على الأقل في الخطوة 2.",
    need_dates: "تاريخ «من» بعد تاريخ «إلى».",
    need_output: "اختر طريقة حفظ واحدة على الأقل في الخطوة 5.",
    need_dir: "اختر مجلدًا لحفظ الأرشيف فيه.",
    nothing_match: "لا توجد عناصر مطابقة لهذه الاختيارات، فلا يوجد ما يُؤرشف.",
    preview_title: "العناصر المطابقة",
    by_type: "حسب النوع",
    by_year: "حسب السنة",
    col_date: "التاريخ",
    col_type: "النوع",
    col_from: "من",
    col_subject: "الموضوع",
    col_folder: "المجلد",
    pv_more: "يُعرض أول {shown} من {total}. قائمة Excel ستحتوي عليها كلها.",
    no_subject: "(بدون موضوع)",
    confirm_title: "جاهز للأرشفة",
    conf_items: "{n} عنصرًا",
    conf_folders: "المجلدات",
    conf_period: "الفترة",
    conf_types: "الأنواع",
    conf_as: "طريقة الحفظ",
    conf_where: "مكان الحفظ",
    conf_skipped: "سبقت أرشفته",
    conf_skipped_v: "تم تخطي {n} عنصرًا",
    what_pst: "ملف PST",
    what_files: "ملفات ({list})",
    what_atts: "المرفقات",
    what_index: "قائمة Excel",
    confirm_move: "بعد فحص النسخة، ستُحذف هذه العناصر من صندوق البريد (تُنقل إلى «العناصر المحذوفة»). جهات الاتصال والمهام والاجتماعات المتكررة والقادمة تبقى.",
    confirm_move_ok: "أفهم أن العناصر ستُحذف من صندوق البريد",
    confirm_copy: "لن يُحذف أي شيء من Outlook. أبقِ Outlook مفتوحًا أثناء العمل.",
    cancel: "إلغاء",
    start: "ابدأ",
    ok: "حسنًا",
    close: "إغلاق",
    job_connect: "جارٍ الاتصال بـ Outlook",
    job_count: "جارٍ البحث عن العناصر المطابقة",
    job_archive: "جارٍ الأرشفة",
    elapsed: "الوقت المنقضي: {t}",
    remaining: "المتبقي تقريبًا: {t}",
    stop: "إيقاف",
    stopping: "جارٍ الإيقاف…",
    stopped: "تم الإيقاف. لم يتم تنفيذ أي شيء آخر.",
    done_title: "الأرشيف جاهز",
    done_stopped: "تم الإيقاف قبل النهاية",
    done_items: "عنصرًا مؤرشفًا",
    done_files: "ملفًا محفوظًا",
    done_removed: "حُذف من صندوق البريد",
    done_kept: "بقي في صندوق البريد",
    done_problems: "مشكلات",
    done_where: "مكان الحفظ:",
    problems_title: "اعرض المشكلات ({n})",
    open_folder: "افتح مجلد الأرشيف",
    sec: "{n} ث",
    min: "{m} د {s} ث",
    hour: "{h} س {m} د",
  },
};

let LANG = "en";
function t(key, p) {
  let s = (STR[LANG] && STR[LANG][key]) ?? STR.en[key] ?? key;
  if (p) s = s.replace(/\{(\w+)\}/g, (m, k) => (k in p ? p[k] : m));
  return s;
}
const fmtN = (n) => Number(n || 0).toLocaleString("en-US");
function el(tag, attrs, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (k === "class") e.className = v;
    else if (k === "text") e.textContent = v;
    else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
    else if (v !== false && v != null) e.setAttribute(k, v === true ? "" : v);
  }
  for (const kid of kids) if (kid != null) e.append(kid);
  return e;
}
function fmtSize(bytes) {
  if (bytes == null) return t("sp_unknown");
  const b = Math.max(0, Number(bytes));
  let v, u;
  if (b >= 1024 ** 3) { v = b / 1024 ** 3; u = "GB"; }
  else if (b >= 1024 ** 2) { v = b / 1024 ** 2; u = "MB"; }
  else { v = Math.max(b / 1024, b ? 1 : 0); u = "KB"; }
  const n = v >= 100 ? Math.round(v).toString() : v.toFixed(1).replace(/\.0$/, "");
  return "\u2066" + n + " " + u + "\u2069"; // stays readable inside Arabic text
}
function fmtTime(sec) {
  sec = Math.max(0, Math.round(sec));
  if (sec < 60) return t("sec", { n: sec });
  if (sec < 3600) return t("min", { m: Math.floor(sec / 60), s: sec % 60 });
  return t("hour", { h: Math.floor(sec / 3600), m: Math.floor((sec % 3600) / 60) });
}
function isoDate(d) {
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

// ----------------------------------------------------------------- server
async function api(path, body) {
  const opts = { headers: { "X-Token": TOKEN } };
  if (body !== undefined) {
    opts.method = "POST";
    opts.headers["Content-Type"] = "application/json";
    opts.body = JSON.stringify(body);
  }
  const r = await fetch("/api/" + path, opts);
  let data = {};
  try { data = await r.json(); } catch (e) { /* empty answer */ }
  if (!r.ok) throw new Error(data.error || "HTTP " + r.status);
  return data;
}

// ------------------------------------------------------------------ state
const KINDS = ["mail", "meeting", "appointment", "contact", "task", "note", "other"];
const KIND_ICON = { mail: "✉️", meeting: "📨", appointment: "📅", contact: "👤", task: "✔️", note: "🗒️", other: "📁" };
const PERIODS = ["all", "this_year", "last_year", "last12", "older1", "older2", "older3", "older5", "year", "custom"];
const DATE_FIELDS = ["auto", "received", "sent", "created", "modified"];
const FORMATS = ["eml", "msg", "mht", "pdf", "txt"];
const LAYOUTS = ["folder_year", "type_year", "year_month", "flat"];

const ui = {
  kinds: new Set(KINDS), period: "all", year: String(new Date().getFullYear() - 1),
  date_from: "", date_to: "", date_field: "auto", attachments: "any", formats: new Set(),
};
const S = {
  server: null, tree: [], treeVersion: -1, checked: new Set(), expanded: new Set(),
  foldersFromSettings: null, seq: { error: 0, result: 0, stopped: 0 },
  scanId: null, scanInfo: null, countedKey: null, pendingKey: null, pendingArchive: false,
  wasBusy: false, jobKind: null, nameTouched: false, singleName: null, lastResult: null, historyCount: 0, ready: false,
};

// ---------------------------------------------------------- folder tree
function subtree(node, out = []) {
  if (!node.is_store) out.push(node.label);
  for (const c of node.children) subtree(c, out);
  return out;
}
function treeState(node) {
  const labels = subtree(node);
  if (!labels.length) return "off";
  const on = labels.filter((l) => S.checked.has(l)).length;
  return on === 0 ? "off" : on === labels.length ? "on" : "some";
}
function nodeMatches(node, q) {
  return !q || node.name.toLowerCase().includes(q) || node.children.some((c) => nodeMatches(c, q));
}
function renderTree() {
  const box = $("tree");
  const q = $("folder-search").value.trim().toLowerCase();
  box.textContent = "";
  if (!S.tree.length) {
    const st = S.server;
    const text = st && st.connected ? t("tree_none") : st && !st.busy ? t("tree_noconn") : t("tree_wait");
    box.append(el("div", { class: "empty", text }));
    return;
  }
  let shown = 0;
  const draw = (node, parent) => {
    if (!nodeMatches(node, q)) return;
    shown++;
    const open = !!q || S.expanded.has(node.key);
    const state = treeState(node);
    const box2 = el("input", { type: "checkbox", tabindex: "-1", "aria-hidden": "true" });
    box2.checked = state === "on";
    box2.indeterminate = state === "some";
    const caret = el("button", {
      type: "button", class: "caret" + (node.children.length ? "" : " leaf") + (open ? " open" : ""),
      "aria-label": "expand", text: open ? "▼" : "▶",
      onclick: (e) => { e.stopPropagation(); S.expanded.has(node.key) ? S.expanded.delete(node.key) : S.expanded.add(node.key); renderTree(); },
    });
    const row = el("div", {
      class: "node-row" + (node.is_store ? " store" : ""), role: "treeitem", tabindex: "0",
      "aria-checked": state === "on" ? "true" : state === "some" ? "mixed" : "false",
      title: node.size != null ? `${node.label} · ${fmtSize(node.size).replace(/[\u2066\u2069]/g, "")}` : node.label,
      onclick: () => toggleNode(node),
      onkeydown: (e) => { if (e.key === " " || e.key === "Enter") { e.preventDefault(); toggleNode(node); } },
    }, caret, box2, el("span", { class: "name", text: node.name }),
      el("span", { class: "count", text: node.is_store ? (node.size != null ? fmtSize(node.size) : "") : node.count >= 0 ? fmtN(node.count) : "" }));
    parent.append(row);
    if (node.children.length && open) {
      const kids = el("div", { class: "children", role: "group" });
      node.children.forEach((c) => draw(c, kids));
      parent.append(kids);
    }
  };
  S.tree.forEach((s) => draw(s, box));
  if (!shown) box.append(el("div", { class: "empty", text: t("tree_nomatch") }));
}
function toggleNode(node) {
  const labels = subtree(node);
  const turnOn = !labels.every((l) => S.checked.has(l));
  labels.forEach((l) => (turnOn ? S.checked.add(l) : S.checked.delete(l)));
  renderTree();
  changed();
}
function setAll(on) {
  S.tree.forEach((s) => subtree(s).forEach((l) => (on ? S.checked.add(l) : S.checked.delete(l))));
  renderTree();
  changed();
}
function checkedFolderKeys() {
  const keys = [];
  const walk = (n) => { if (!n.is_store && S.checked.has(n.label)) keys.push(n.key); n.children.forEach(walk); };
  S.tree.forEach(walk);
  return keys;
}
async function loadTree() {
  const data = await api("tree");
  S.tree = data.tree;
  S.treeVersion = data.version;
  const all = new Set();
  S.tree.forEach((s) => subtree(s).forEach((l) => all.add(l)));
  if (S.foldersFromSettings) {
    S.checked = new Set(S.foldersFromSettings.filter((l) => all.has(l)));
    S.foldersFromSettings = null;
  } else if (!S.checked.size && S.tree.length) {
    subtree(S.tree[0]).forEach((l) => S.checked.add(l));
  }
  S.tree.forEach((s) => S.expanded.add(s.key));
  renderTree();
  updateSummary();
}

// ---------------------------------------------------------------- choices
function chip(label, pressed, onclick, icon, extra) {
  return el("button", { type: "button", class: "chip", "aria-pressed": pressed ? "true" : "false", onclick },
    icon ? el("span", { class: "ico", text: icon }) : null, el("span", { text: label }),
    extra ? el("bdi", { class: "ext", text: extra }) : null);
}
const FORMAT_EXT = { eml: ".eml .ics .vcf", msg: ".msg", mht: ".mht", txt: ".txt" };
function caps() {
  const c = (S.server && S.server.capabilities) || {};
  return { pst: c.pst !== false, kinds: c.kinds || KINDS, formats: c.formats || ["msg", "mht", "pdf", "txt"] };
}
function renderChips() {
  const kinds = $("kinds");
  kinds.textContent = "";
  const cap = caps();
  KINDS.filter((k) => cap.kinds.includes(k)).forEach((k) => kinds.append(chip(t("kind_" + k), ui.kinds.has(k), () => {
    ui.kinds.has(k) ? ui.kinds.delete(k) : ui.kinds.add(k);
    renderChips(); changed();
  }, KIND_ICON[k])));

  const per = $("periods");
  per.textContent = "";
  PERIODS.forEach((p) => per.append(chip(t("p_" + p), ui.period === p, () => {
    ui.period = p; renderChips(); renderPeriod(); changed(); autoName();
  })));

  const fm = $("formats");
  fm.textContent = "";
  FORMATS.filter((f) => cap.formats.includes(f)).forEach((f) => fm.append(chip(t("fmt_" + f), ui.formats.has(f), () => {
    ui.formats.has(f) ? ui.formats.delete(f) : ui.formats.add(f);
    renderChips(); changed();
  }, null, f === "mht" && !cap.pst ? ".html" : FORMAT_EXT[f] || t("fmt_pdf_note"))));

  const seg = $("attachments");
  seg.textContent = "";
  ["any", "with", "without"].forEach((a) => seg.append(el("button", {
    type: "button", role: "radio", "aria-checked": ui.attachments === a ? "true" : "false", text: t("att_" + a),
    onclick: () => { ui.attachments = a; renderChips(); changed(); },
  })));
}
function fillSelect(sel, values, label, current) {
  sel.textContent = "";
  values.forEach((v) => sel.append(el("option", { value: v, text: label(v) })));
  if (current != null) sel.value = current;
}
function renderSelects() {
  const y = new Date().getFullYear();
  const years = Array.from({ length: 30 }, (_, i) => String(y - i));
  fillSelect($("year"), years, (v) => v, ui.year);
  fillSelect($("date-field"), DATE_FIELDS, (v) => t("df_" + v), ui.date_field);
  fillSelect($("layout"), LAYOUTS, (v) => t("lay_" + v), $("layout").value || "folder_year");
}
function periodRange() {
  const now = new Date();
  const y = now.getFullYear();
  const older = (n) => { const d = new Date(y - n, now.getMonth(), now.getDate()); d.setDate(d.getDate() - 1); return isoDate(d); };
  switch (ui.period) {
    case "this_year": return [`${y}-01-01`, ""];
    case "last_year": return [`${y - 1}-01-01`, `${y - 1}-12-31`];
    case "last12": { const d = new Date(now); d.setFullYear(y - 1); return [isoDate(d), ""]; }
    case "older1": return ["", older(1)];
    case "older2": return ["", older(2)];
    case "older3": return ["", older(3)];
    case "older5": return ["", older(5)];
    case "year": return [`${ui.year}-01-01`, `${ui.year}-12-31`];
    case "custom": return [$("date-from").value, $("date-to").value];
    default: return ["", ""];
  }
}
const ltr = (s) => (s ? "\u2066" + s + "\u2069" : s); // keep dates left-to-right inside Arabic text
function rangeText() {
  const [a0, b0] = periodRange();
  const a = ltr(a0), b = ltr(b0);
  if (a && b) return t("range_both", { a, b });
  if (b) return t("range_until", { b });
  if (a) return t("range_since", { a });
  return t("range_all");
}
function renderPeriod() {
  $("year").hidden = ui.period !== "year";
  $("custom").hidden = ui.period !== "custom";
  $("range").textContent = rangeText();
}
function filterValues() {
  return {
    sender: $("f-sender").value.trim(), recipients: $("f-recipients").value.trim(),
    subject: $("f-subject").value.trim(), category: $("f-category").value.trim(),
    body_words: $("f-body").value.trim(),
  };
}
function activeFilters() {
  const n = Object.values(filterValues()).filter(Boolean).length;
  return n + (ui.attachments !== "any" ? 1 : 0);
}
function renderFilterBadge() {
  const n = activeFilters();
  const b = $("filters-badge");
  b.textContent = n ? t("s4_active", { n }) : t("optional");
  b.className = "badge" + (n ? "" : " muted");
}
function autoName() {
  if (S.nameTouched) return;
  // one email chosen: the archive is named after it
  if (S.singleName && S.scanInfo && S.countedKey === JSON.stringify(countRequest())) {
    $("o-name").value = S.singleName;
    return;
  }
  const [a, b] = periodRange();
  let suffix;
  if (a && b && a.slice(0, 4) === b.slice(0, 4) && a.endsWith("-01-01") && b.endsWith("-12-31")) suffix = a.slice(0, 4);
  else if (a && b) suffix = `${a} - ${b}`;
  else if (b) suffix = t("name_until", { d: b });
  else if (a) suffix = t("name_since", { d: a });
  else suffix = isoDate(new Date());
  $("o-name").value = `${t("name_base")} ${suffix}`;
}
function criteria() {
  const [date_from, date_to] = periodRange();
  return { kinds: KINDS.filter((k) => ui.kinds.has(k) && caps().kinds.includes(k)), date_from, date_to,
    date_field: $("date-field").value, attachments: ui.attachments, ...filterValues() };
}
function countRequest() {
  return { folders: checkedFolderKeys(), criteria: criteria(), skip_archived: $("o-skip").checked };
}
function exportOptions() {
  return {
    output_dir: $("o-dir").value.trim(), archive_name: $("o-name").value.trim(),
    make_pst: caps().pst && $("o-pst").checked, keep_pst_open: $("o-pst-show").checked,
    file_formats: FORMATS.filter((f) => ui.formats.has(f) && caps().formats.includes(f)), save_attachments: $("o-atts").checked,
    write_index: $("o-index").checked, layout: $("layout").value, move_after: $("o-move").checked,
  };
}
function collectSettings() {
  const [from, to] = periodRange();
  return {
    lang: LANG, kinds: [...ui.kinds], period: ui.period, year: ui.year,
    date_from: ui.period === "custom" ? from : "", date_to: ui.period === "custom" ? to : "",
    date_field: $("date-field").value, attachments: ui.attachments, ...filterValues(),
    make_pst: $("o-pst").checked, keep_pst_open: $("o-pst-show").checked,
    file_formats: [...ui.formats], save_attachments: $("o-atts").checked, write_index: $("o-index").checked,
    layout: $("layout").value, output_dir: $("o-dir").value.trim(), skip_archived: $("o-skip").checked,
    folders: [...S.checked],
  };
}
function applySettings(st) {
  ui.kinds = new Set((st.kinds || KINDS).filter((k) => KINDS.includes(k)));
  ui.period = PERIODS.includes(st.period) ? st.period : "all";
  if (st.year) ui.year = String(st.year);
  ui.date_field = DATE_FIELDS.includes(st.date_field) ? st.date_field : "auto";
  ui.attachments = ["any", "with", "without"].includes(st.attachments) ? st.attachments : "any";
  ui.formats = new Set((st.file_formats || []).filter((f) => FORMATS.includes(f)));
  $("date-from").value = st.date_from || "";
  $("date-to").value = st.date_to || "";
  $("f-sender").value = st.sender || "";
  $("f-recipients").value = st.recipients || "";
  $("f-subject").value = st.subject || "";
  $("f-category").value = st.category || "";
  $("f-body").value = st.body_words || "";
  $("o-pst").checked = st.make_pst !== false;
  $("o-pst-show").checked = st.keep_pst_open !== false;
  $("o-atts").checked = !!st.save_attachments;
  $("o-index").checked = st.write_index !== false;
  $("o-skip").checked = st.skip_archived !== false;
  $("o-dir").value = st.output_dir || "";
  $("o-move").checked = false; // never remembered, on purpose
  renderSelects();
  $("layout").value = LAYOUTS.includes(st.layout) ? st.layout : "folder_year";
  if (Array.isArray(st.folders)) S.foldersFromSettings = st.folders;
  if (activeFilters()) $("filters-card").open = true;
}

// -------------------------------------------------------------- language
function applyLang() {
  document.documentElement.lang = LANG;
  document.documentElement.dir = LANG === "ar" ? "rtl" : "ltr";
  document.title = t("app_title");
  document.querySelectorAll("[data-t]").forEach((e) => (e.textContent = t(e.dataset.t)));
  document.querySelectorAll("[data-tp]").forEach((e) => (e.placeholder = t(e.dataset.tp)));
  $("lang-btn").textContent = t("lang_other");
  renderSelects();
  renderChips();
  renderPeriod();
  renderFilterBadge();
  renderTree();
  renderHistory();
  autoName();
  updateSummary();
  updatePill();
}

// ------------------------------------------------------- summary & buttons
let saveTimer = null;
function saveNow() {
  clearTimeout(saveTimer);
  saveTimer = null;
  return api("settings", collectSettings()).catch(() => {});
}
function changed(immediately) {
  renderFilterBadge();
  updateSummary();
  updateMoveTile();
  clearTimeout(saveTimer);
  if (immediately === true) saveNow();
  else saveTimer = setTimeout(saveNow, 400);
}
function updateMoveTile() {
  const cap = caps();
  $("pst-tile").hidden = !cap.pst;
  const o = exportOptions();
  const safeCopy = o.make_pst || o.file_formats.includes("eml");
  if (!safeCopy) $("o-move").checked = false;
  $("o-move").disabled = !safeCopy;
  $("move-needs-pst").hidden = safeCopy;
  $("move-needs-pst").textContent = cap.pst ? t("o_move_needs_pst") : t("o_move_needs_eml");
  $("move-tile").classList.toggle("on", $("o-move").checked);
  $("o-pst-show").disabled = !o.make_pst;
}
function updateSummary() {
  const folders = checkedFolderKeys().length;
  const line = $("summary-line");
  if (!folders) {
    line.textContent = t("sum_none");
  } else {
    const types = ui.kinds.size === KINDS.length ? t("sum_all_types")
      : ui.kinds.size <= 3 ? KINDS.filter((k) => ui.kinds.has(k)).map((k) => t("kind_" + k)).join(LANG === "ar" ? "، " : ", ")
        : t("sum_types", { n: ui.kinds.size });
    const parts = [t("sum_folders", { n: folders }), types, rangeText()];
    const f = activeFilters();
    if (f) parts.push(t("sum_filters", { n: f }));
    line.textContent = parts.join("  ·  ");
  }
  const cl = $("count-line");
  cl.textContent = "";
  cl.className = "count-line";
  const key = JSON.stringify(countRequest());
  if (S.scanInfo && S.countedKey === key) {
    const total = S.scanInfo.summary.total;
    cl.append(el("strong", { text: total === 1 ? t("count_one") : t("count_res", { n: fmtN(total) }) }));
    if (total) cl.append("  ·  " + t("count_size", { s: fmtSize(S.scanInfo.summary.size) }));
    if (S.scanInfo.skipped) cl.append("  ·  " + t("count_skipped", { n: fmtN(S.scanInfo.skipped) }));
    if (S.scanInfo.summary.total) cl.append("  ·  ", el("button", { type: "button", class: "link small", text: t("show_list"), onclick: openPreview }));
  } else if (S.scanInfo) {
    cl.textContent = t("count_stale");
    cl.classList.add("stale");
  } else {
    cl.textContent = t("count_none");
  }
  const st = S.server;
  const ready = st && st.connected && !st.busy;
  $("btn-count").disabled = !ready || !folders;
  $("btn-archive").disabled = !ready || !folders;
}
function renderHistory() {
  $("history-text").textContent = S.historyCount ? t("hist_some", { n: fmtN(S.historyCount) }) : t("hist_none");
  $("forget").hidden = !S.historyCount;
}
function updatePill() {
  const st = S.server;
  const pill = $("conn");
  let cls = "wait", text = t("conn_wait");
  if (st && st.connected) {
    cls = st.demo ? "demo" : "ok";
    text = st.engine === "online" ? t("conn_online", { who: st.account })
      : st.demo ? t("conn_demo") : st.background ? t("conn_bg", { n: st.store_count }) : t("conn_ok", { n: st.store_count });
  } else if (st && !(st.busy && st.job && st.job.kind === "connect") && S.ready) {
    cls = "bad"; text = t("conn_bad");
  }
  pill.className = "pill " + cls;
  $("conn-text").textContent = text;
}

// ---------------------------------------------------------------- dialogs
function message(text, withCancel = false) {
  return new Promise((resolve) => {
    const d = $("dlg-message");
    $("msg-text").textContent = text;
    $("msg-no").hidden = !withCancel;
    const done = (v) => { d.close(); resolve(v); };
    $("msg-yes").onclick = () => done(true);
    $("msg-no").onclick = () => done(false);
    d.oncancel = () => resolve(false);
    d.showModal();
  });
}
function showBanner(text, retry) {
  $("banner").classList.remove("info");
  $("banner-text").textContent = text;
  $("banner-retry").hidden = !retry;
  const st = S.server || {};
  $("banner-online").hidden = !(retry && st.classic_failed && (st.engines || []).includes("online"));
  $("banner").hidden = false;
}
async function openConnection() {
  const st = S.server || {};
  let saved = {};
  try { saved = await api("settings"); } catch (e) { /* keep the fields as they are */ }
  $("client-id").value = saved.client_id || "";
  $("tenant").value = saved.tenant && saved.tenant !== "organizations" ? saved.tenant : "";
  const engines = st.engines || [];
  const set = (id, eng) => {
    const box = $("engine-" + eng);
    const inUse = st.connected && st.engine === eng;
    box.classList.toggle("current", inUse);
    const badge = $(id);
    badge.textContent = inUse ? t("eng_in_use") : engines.includes(eng) ? "" : t("eng_unavailable");
    badge.className = "badge" + (inUse ? "" : " muted");
  };
  set("eng-classic-state", "classic");
  set("eng-online-state", "online");
  $("use-classic").disabled = !engines.includes("classic") || st.busy;
  $("use-online").disabled = !engines.includes("online") || st.busy;
  $("sign-out").hidden = !(st.connected && st.engine === "online");
  $("engine-online").querySelectorAll(".field, .advanced, .hint:not(:first-of-type)").forEach((e) => (e.hidden = !st.needs_client_id));
  if (!$("dlg-conn").open) $("dlg-conn").showModal();
}
async function useEngine(engine) {
  if (engine === "online" && (S.server || {}).needs_client_id) {
    const id = $("client-id").value.trim();
    if (!id) { $("client-id").focus(); return message(t("need_client_id")); }
    await api("settings", { client_id: id, tenant: $("tenant").value.trim() || "organizations" }).catch(() => {});
  }
  $("dlg-conn").close();
  $("banner").hidden = true;
  S.checked = new Set(); // folders differ between the two connections
  S.foldersFromSettings = null;
  try { await api("connect", { engine }); } catch (e) { return message(e.message); }
  poll();
}
function validateChoices(forArchive) {
  if (!checkedFolderKeys().length) return t("need_folders");
  if (!ui.kinds.size) return t("need_kinds");
  const [a, b] = periodRange();
  if (a && b && a > b) return t("need_dates");
  if (forArchive) {
    const o = exportOptions();
    if (!(o.make_pst || o.file_formats.length || o.save_attachments || o.write_index)) return t("need_output");
    if (!o.output_dir) return t("need_dir");
  }
  return null;
}
async function startCount(thenArchive) {
  const problem = validateChoices(thenArchive);
  if (problem) return message(problem);
  const req = countRequest();
  S.pendingKey = JSON.stringify(req);
  S.pendingArchive = thenArchive;
  try {
    await api("count", req);
    S.jobKind = "count";
    openProgress("count");
    poll();
  } catch (e) {
    S.pendingArchive = false;
    message(e.message);
  }
}
function onArchiveClick() {
  const problem = validateChoices(true);
  if (problem) return message(problem);
  if (S.scanInfo && S.countedKey === JSON.stringify(countRequest())) {
    if (!S.scanInfo.summary.total) return message(t("nothing_match"));
    openConfirm();
  } else {
    startCount(true);
  }
}
function openConfirm() {
  const o = exportOptions();
  const info = S.scanInfo;
  $("conf-number").textContent = info.summary.total === 1 ? t("conf_one") : t("conf_items", { n: fmtN(info.summary.total) });
  const facts = $("conf-facts");
  facts.textContent = "";
  const what = [];
  if (o.make_pst) what.push(t("what_pst"));
  if (o.file_formats.length) what.push(t("what_files", { list: o.file_formats.map((f) => f.toUpperCase()).join(", ") }));
  if (o.save_attachments) what.push(t("what_atts"));
  if (o.write_index) what.push(t("what_index"));
  const types = ui.kinds.size === KINDS.length ? t("sum_all_types") : [...ui.kinds].map((k) => t("kind_" + k)).join(LANG === "ar" ? "، " : ", ");
  const sep = o.output_dir.includes("/") && !o.output_dir.includes("\\") ? "/" : "\\";
  const rows = [
    ["conf_folders", fmtN(checkedFolderKeys().length)],
    ["conf_types", types],
    ["conf_period", rangeText()],
    ["conf_as", what.join(LANG === "ar" ? "، " : ", ")],
    ["conf_where", o.output_dir + sep + o.archive_name],
  ];
  if (info.skipped) rows.push(["conf_skipped", t("conf_skipped_v", { n: fmtN(info.skipped) })]);
  rows.forEach(([k, v]) => {
    const dd = el("dd");
    if (k === "conf_where") dd.append(el("bdi", { dir: "ltr", text: v })); else dd.textContent = v;
    facts.append(el("dt", { text: t(k) }), dd);
  });
  $("conf-move").hidden = !o.move_after;
  $("conf-copy").hidden = o.move_after;
  $("conf-move-ok").checked = false;
  const start = $("conf-start");
  start.classList.toggle("danger", o.move_after);
  S.spaceOk = false;
  updateStart();
  const space = $("conf-space");
  space.textContent = "";
  space.append(el("h4", { text: t("space_title") }), el("p", { class: "hint", text: t("space_calc") }));
  $("dlg-confirm").showModal();
  loadSpace(o);
}
function updateStart() {
  const moveOk = !exportOptions().move_after || $("conf-move-ok").checked;
  $("conf-start").disabled = !(S.spaceOk && moveOk);
}
async function loadSpace(o) {
  const space = $("conf-space");
  let e;
  try {
    e = await api("estimate", { scan_id: S.scanId, options: o });
  } catch (err) {
    space.lastChild.textContent = err.message;
    S.spaceOk = true; // the numbers are a guide; do not block on them
    updateStart();
    return;
  }
  space.textContent = "";
  space.append(el("h4", { text: t("space_title") }));
  const table = el("div", { class: "space-grid" });
  const head = (a, b, c) => table.append(el("span", { class: "sh", text: a }), el("span", { class: "sh", text: b }), el("span", { class: "sh", text: c }));
  const row = (label, now, after, note, cls) => {
    table.append(el("span", { class: "sl" }, label, note ? el("small", { text: note }) : null),
      el("span", { class: "sv", text: now }), el("span", { class: "sv " + (cls || ""), text: after }));
  };
  head("", t("sp_now"), t("sp_after"));
  row(t("sp_items"), fmtSize(e.items_size), "", t("sp_items_n", { n: fmtN(S.scanInfo.summary.total) }), "muted");
  const freeAfter = e.free == null ? null : e.free - e.archive;
  row(e.drive ? t("sp_disk_d", { d: e.drive }) : t("sp_disk"), fmtSize(e.free),
    e.free == null ? "" : t("approx", { s: fmtSize(Math.max(0, freeAfter)) }),
    t("sp_archive", { s: fmtSize(e.archive) }), e.enough ? "" : "bad");
  e.mailboxes.forEach((m) => {
    const after = e.move && m.size != null ? t("approx", { s: fmtSize(Math.max(0, m.size - m.removed)) }) : e.move ? "" : t("sp_nochange");
    row(t("sp_mailbox", { n: m.name }), fmtSize(m.size), after,
      t("sp_items_v", { s: fmtSize(m.chosen) }) + (e.move ? " · " + t("sp_after_empty") : ""), e.move ? "good" : "muted");
    if (m.file_size != null) {
      row(t("sp_file", { ext: m.file_ext || ".ost" }), fmtSize(m.file_size),
        e.move ? t("sp_file_shrink") : t("sp_nochange"), "", "muted");
    }
  });
  space.append(table);
  if (!e.enough) {
    space.append(el("p", { class: "callout danger", text: t("sp_not_enough", { need: fmtSize(e.archive), free: fmtSize(e.free) }) }));
  }
  S.spaceOk = e.enough;
  updateStart();
}
async function startArchive() {
  $("dlg-confirm").close();
  try {
    await api("archive", { scan_id: S.scanId, options: exportOptions() });
    S.jobKind = "archive";
    openProgress("archive");
    poll();
  } catch (e) {
    message(e.message);
  }
}
function openProgress(kind) {
  $("prog-title").textContent = t("job_" + kind);
  $("prog-text").textContent = "";
  $("prog-time").textContent = "";
  $("prog-bar").style.width = "0";
  $("prog-stop").disabled = false;
  $("prog-stop").textContent = t("stop");
  const d = $("dlg-progress");
  if (!d.open) d.showModal();
}
function updateProgress(job) {
  const d = $("dlg-progress");
  if (!d.open || !job) return;
  $("prog-title").textContent = t("job_" + job.kind);
  const bar = d.querySelector(".progress");
  if (job.total > 0) {
    bar.classList.remove("indeterminate");
    $("prog-bar").style.width = Math.min(100, (100 * job.done) / job.total).toFixed(1) + "%";
  } else {
    bar.classList.add("indeterminate");
  }
  $("prog-text").textContent = job.text || "";
  let time = t("elapsed", { t: fmtTime(job.elapsed) });
  if (job.kind === "archive" && job.total > 0 && job.done >= 5 && job.done < job.total) {
    time += "  ·  " + t("remaining", { t: fmtTime((job.elapsed / job.done) * (job.total - job.done)) });
  }
  $("prog-time").textContent = time;
}
function showDone(r) {
  S.lastResult = r;
  $("o-move").checked = false; // ask again each time, on purpose
  updateMoveTile();
  const removedAsked = r.removed > 0 || r.kept > 0;
  const c = r.check || {};
  const bad = r.cancelled || c.ok === false;
  $("done-title").textContent = r.cancelled ? t("done_stopped") : c.ok === false ? t("done_title_bad") : t("done_title");
  $("done-icon").textContent = bad ? "!" : "✓";
  $("done-icon").className = "done-icon" + (bad ? " warn" : "");
  renderCheck(c);
  const tiles = $("done-tiles");
  tiles.textContent = "";
  const stat = (n, label, bad, text) => tiles.append(el("div", { class: "stat" + (bad ? " bad" : "") }, el("b", { text: text || fmtN(n) }), el("span", { text: label })));
  stat(r.total, t("done_items"));
  if (r.files_written) stat(r.files_written, t("done_files"));
  stat(0, t("done_size"), false, fmtSize(r.archive_size));
  if (removedAsked) stat(r.removed, t("done_removed"));
  if (r.removed_size) stat(0, t("done_freed"), false, fmtSize(r.removed_size));
  if (r.kept) stat(r.kept, t("done_kept"));
  if (r.errors.length) stat(r.errors.length, t("done_problems"), true);
  $("done-folder").textContent = r.folder;
  const pr = $("done-problems");
  pr.hidden = !r.errors.length;
  pr.open = false;
  $("done-problems-title").textContent = t("problems_title", { n: r.errors.length });
  const list = $("done-problems-list");
  list.textContent = "";
  r.errors.slice(0, 300).forEach((e) => list.append(el("li", { text: e })));
  $("dlg-done").showModal();
}
function renderCheck(c) {
  const box = $("done-check");
  box.textContent = "";
  if (!c || c.chosen == null) { box.hidden = true; return; }
  box.hidden = false;
  box.className = "check " + (c.ok ? "good" : "bad");
  box.append(el("h4", { text: t("chk_title") }));
  const line = (ok, text) => box.append(el("div", { class: "check-line" },
    el("span", { class: "mark " + (ok ? "ok" : "no"), text: ok ? "✓" : "✗" }), el("span", { text })));
  line(c.ok_items === c.chosen, t("chk_items", { a: fmtN(c.chosen), b: fmtN(c.ok_items) }));
  if (c.pst) {
    line(c.pst.checked && c.pst.found === c.pst.expected, t("chk_pst", {
      a: fmtN(c.pst.found), b: fmtN(c.pst.expected), s1: fmtSize(c.pst.expected_size), s2: fmtSize(c.pst.found_size) }));
  }
  if (c.files) line(c.files.found >= c.files.expected, t("chk_files", { a: fmtN(c.files.found), b: fmtN(c.files.expected) }));
  box.append(el("p", { class: "check-result", text: c.ok ? t("chk_all_ok") : t("chk_bad", { n: fmtN(c.missing_count) }) }));
  if (c.pst && c.pst.found === c.pst.expected && c.pst.expected_size && Math.abs(c.pst.found_size - c.pst.expected_size) > 0.02 * c.pst.expected_size) {
    box.append(el("p", { class: "hint", text: t("chk_size_note") }));
  }
  if (c.missing && c.missing.length) {
    const list = el("ul");
    c.missing.slice(0, 200).forEach((m) => list.append(el("li", {},
      el("bdi", { dir: "ltr", text: m.date }), "  ", el("b", { text: m.subject || t("no_subject") }), "  · " + m.reason)));
    box.append(el("details", {}, el("summary", { text: t("chk_show", { n: fmtN(c.missing_count) }) }), list));
  }
}
async function openPreview() {
  const data = await api("scan");
  const info = S.scanInfo;
  if (!info) return;
  const types = $("pv-types");
  types.textContent = "";
  KINDS.filter((k) => info.summary.by_kind[k]).forEach((k) =>
    types.append(el("span", { class: "chip" }, el("span", { class: "ico", text: KIND_ICON[k] }), t("kind_" + k) + " ", el("b", { text: fmtN(info.summary.by_kind[k]) }))));
  const years = $("pv-years");
  years.textContent = "";
  const entries = Object.entries(info.summary.by_year);
  const max = Math.max(1, ...entries.map(([, n]) => n));
  entries.forEach(([y, n]) => years.append(el("div", { class: "bar-row" },
    el("span", { text: y === "-" ? "—" : y }),
    el("span", { class: "track" }, el("span", { class: "fill", style: `width:${(100 * n) / max}%` })),
    el("span", { class: "n", text: fmtN(n) }))));
  const body = $("pv-rows");
  body.textContent = "";
  data.rows.forEach((r) => body.append(el("tr", {},
    el("td", { dir: "ltr", text: r.date }), el("td", { text: t("kind_" + r.kind) }), el("td", { text: r.from }),
    el("td", { class: "subject", text: r.subject || t("no_subject") }), el("td", { class: "num", text: fmtSize(r.size) }),
    el("td", { class: "folder", text: r.folder }))));
  $("pv-total").textContent = t("pv_total", { s: fmtSize(info.summary.size) });
  $("pv-more").textContent = data.more ? t("pv_more", { shown: fmtN(data.rows.length), total: fmtN(info.summary.total) }) : "";
  $("dlg-preview").showModal();
}

async function nameFromScan(scan) {
  S.singleName = null;
  if (scan.summary.total === 1) {
    try {
      const row = (await api("scan")).rows[0] || {};
      const subject = (row.subject || "").replace(/\s+/g, " ").trim();
      S.singleName = (subject || `${t("no_subject")} ${(row.date || "").slice(0, 10)}`.trim()).slice(0, 80).trim();
    } catch (e) { /* keep the usual name */ }
  }
  autoName();
}

// ----------------------------------------------------------------- polling
let pollTimer = null;
async function poll() {
  clearTimeout(pollTimer);
  try {
    const st = await api("state");
    await onState(st);
  } catch (e) { /* the program may be starting or closing */ }
  const fast = S.server && S.server.busy;
  pollTimer = setTimeout(poll, fast ? 400 : 1500);
}
async function onState(st) {
  const prev = S.server;
  S.server = st;
  S.ready = true;
  if (st.tree_version !== S.treeVersion && st.tree_version > 0) await loadTree();
  if (st.history_count !== S.historyCount) { S.historyCount = st.history_count; renderHistory(); }

  if (st.busy && st.job) {
    if (st.job.kind !== "connect") { if (!$("dlg-progress").open) openProgress(st.job.kind); updateProgress(st.job); }
  } else if ($("dlg-progress").open) {
    $("dlg-progress").close();
  }

  // a new count finished
  const scanId = st.scan ? st.scan.id : null;
  if (scanId !== S.scanId) {
    S.scanId = scanId;
    S.scanInfo = st.scan;
    if (st.scan) {
      S.countedKey = S.pendingKey;
      await nameFromScan(st.scan);
      if (S.pendingArchive) {
        S.pendingArchive = false;
        if (!st.scan.summary.total) message(t("nothing_match")); else openConfirm();
      }
    }
  }
  if (st.error_seq !== S.seq.error) {
    S.seq.error = st.error_seq;
    S.pendingArchive = false;
    if (st.error) showBanner(st.error, !st.connected);
  }
  if (st.stopped_seq !== S.seq.stopped) {
    if (S.seq.stopped !== -1) message(t("stopped"));
    S.seq.stopped = st.stopped_seq;
    S.pendingArchive = false;
  }
  if (st.result_seq !== S.seq.result) {
    const first = S.seq.result === -1;
    S.seq.result = st.result_seq;
    if (!first && st.result) {
      showDone(st.result);
      if (st.result.removed) api("connect", {}).catch(() => {}); // refresh folder counts
    }
  }
  if (st.connected && !$("banner").hidden && !$("banner-retry").hidden) $("banner").hidden = true;
  if (!prev || prev.busy !== st.busy || prev.connected !== st.connected) { updatePill(); if (!S.tree.length) renderTree(); }
  if (!prev || prev.engine !== st.engine || prev.connected !== st.connected) { renderChips(); updateMoveTile(); }
  if (st.background && !S.bgShown) {
    S.bgShown = true;
    $("banner-text").textContent = t("bg_note");
    $("banner-retry").hidden = true;
    $("banner-online").hidden = true;
    $("banner").classList.add("info");
    $("banner").hidden = false;
  }
  updateSummary();
}

// ------------------------------------------------------------------ wiring
function wire() {
  $("lang-btn").onclick = () => { LANG = LANG === "en" ? "ar" : "en"; applyLang(); changed(true); };
  document.querySelectorAll("#dlg-lang [data-lang]").forEach((b) => (b.onclick = () => {
    LANG = b.dataset.lang; $("dlg-lang").close(); applyLang(); changed(true);
  }));
  $("dlg-lang").addEventListener("cancel", (e) => e.preventDefault());
  $("dlg-progress").addEventListener("cancel", (e) => e.preventDefault());
  $("folder-search").addEventListener("input", renderTree);
  $("tick-all").onclick = () => setAll(true);
  $("untick-all").onclick = () => setAll(false);
  const reconnect = () => { $("banner").hidden = true; api("connect", {}).then(poll).catch((e) => message(e.message)); };
  $("reconnect").onclick = reconnect;
  $("banner-retry").onclick = reconnect;
  $("banner-online").onclick = () => { $("banner").hidden = true; openConnection(); };
  $("conn").onclick = openConnection;
  $("conn-close").onclick = () => $("dlg-conn").close();
  $("use-classic").onclick = () => useEngine("classic");
  $("use-online").onclick = () => useEngine("online");
  $("sign-out").onclick = async () => {
    $("dlg-conn").close();
    try { await api("signout", {}); } catch (e) { message(e.message); }
    poll();
  };
  $("banner-close").onclick = () => ($("banner").hidden = true);
  $("year").onchange = () => { ui.year = $("year").value; renderPeriod(); changed(); autoName(); };
  $("date-from").onchange = $("date-to").onchange = () => { renderPeriod(); changed(); autoName(); };
  $("date-field").onchange = changed;
  $("layout").onchange = changed;
  ["f-sender", "f-recipients", "f-subject", "f-category", "f-body"].forEach((id) => $(id).addEventListener("input", changed));
  ["o-pst", "o-pst-show", "o-atts", "o-index", "o-skip", "o-move"].forEach((id) => $(id).addEventListener("change", changed));
  $("o-dir").addEventListener("input", changed);
  $("o-name").addEventListener("input", () => { S.nameTouched = !!$("o-name").value.trim(); });
  $("browse").onclick = async () => {
    try { const r = await api("browse", { current: $("o-dir").value }); if (r.path) { $("o-dir").value = r.path; changed(); } }
    catch (e) { message(e.message); }
  };
  $("forget").onclick = async () => {
    if (await message(t("forget_q"), true)) { await api("forget", {}); S.scanInfo = null; poll(); }
  };
  $("btn-count").onclick = () => startCount(false);
  $("btn-archive").onclick = onArchiveClick;
  $("conf-cancel").onclick = () => $("dlg-confirm").close();
  $("conf-start").onclick = startArchive;
  $("conf-move-ok").onchange = updateStart;
  $("prog-stop").onclick = () => { $("prog-stop").disabled = true; $("prog-stop").textContent = t("stopping"); api("stop", {}).catch(() => {}); };
  $("done-close").onclick = () => $("dlg-done").close();
  $("done-open").onclick = () => S.lastResult && api("open", { path: S.lastResult.folder }).catch((e) => message(e.message));
  $("preview-close").onclick = () => $("dlg-preview").close();
  window.addEventListener("pagehide", () => {
    if (saveTimer) {
      fetch("/api/settings", { method: "POST", keepalive: true,
        headers: { "X-Token": TOKEN, "Content-Type": "application/json" }, body: JSON.stringify(collectSettings()) }).catch(() => {});
    }
    fetch("/api/bye", { method: "POST", keepalive: true, headers: { "X-Token": TOKEN } }).catch(() => {});
  });
}

async function init() {
  wire();
  S.seq = { error: -1, result: -1, stopped: -1 };
  let st = {};
  try { st = await api("settings"); } catch (e) { /* defaults */ }
  LANG = st.lang === "ar" ? "ar" : "en";
  applySettings(st);
  applyLang();
  updateMoveTile();
  if (!st.lang) $("dlg-lang").showModal();
  // remember the counters already on the server so old news is not shown again
  try {
    const first = await api("state");
    S.seq = { error: -1, result: first.result_seq, stopped: first.stopped_seq };
  } catch (e) { /* first poll will catch up */ }
  poll();
}
init();
