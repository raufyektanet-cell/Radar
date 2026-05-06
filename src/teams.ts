export interface TeamDef {
  teamFa: string;
  teamEn: string;
  managerFa: string;
  managerEn: string;
  username: string;
  dept: string;
}

export const TEAMS: TeamDef[] = [
  { teamFa: "منیجد سرویس",       teamEn: "Managed Service",     managerFa: "نرگس ایمانی",         managerEn: "Narges Imani",          username: "narges",    dept: "بیزینس" },
  { teamFa: "موفقیت مشتریان A",  teamEn: "Customer Success A",  managerFa: "نازنین افخمی اوغانی",  managerEn: "Nazanin Afkhami",        username: "nazanin",   dept: "بیزینس" },
  { teamFa: "موفقیت مشتریان B",  teamEn: "Customer Success B",  managerFa: "حمیدرضا رهبر",         managerEn: "Hamidreza Rahbar",       username: "shahab",    dept: "بیزینس" },
  { teamFa: "تعالی مشتریان",     teamEn: "Customer Excellence", managerFa: "مریم لواف",            managerEn: "Maryam Lavvaf",          username: "maryam",    dept: "بیزینس" },
  { teamFa: "پیامک",             teamEn: "SMS",                 managerFa: "فرهاد حیدریان",        managerEn: "Farhad Heydarian",       username: "farhad",    dept: "بیزینس" },
  { teamFa: "ویدئو",             teamEn: "Video",               managerFa: "حامد زمانی",           managerEn: "Hamed Zamani",           username: "hamed",     dept: "بیزینس" },
  { teamFa: "اورانوس",           teamEn: "Uranus",              managerFa: "مهران پاکند",          managerEn: "Mehran Pakand",          username: "mehran",    dept: "محصول" },
  { teamFa: "روبیکا",            teamEn: "Rubika",              managerFa: "رضا میرباقری",         managerEn: "Reza Mirbaghri",         username: "mirbaghri", dept: "محصول" },
  { teamFa: "پنجره",             teamEn: "Panjereh",            managerFa: "محمد زابلیان",         managerEn: "Mohammad Zabolian",      username: "zabol",     dept: "محصول" },
  { teamFa: "ادپلاس",            teamEn: "AdPlus",              managerFa: "میلاد کلانتری",        managerEn: "Milad Kalantari",        username: "milad",     dept: "فروش" },
  { teamFa: "ادوی",              teamEn: "AdWay",               managerFa: "مهدی مظفری",           managerEn: "Mehdi Mozaffari",        username: "mehdi",     dept: "فروش" },
  { teamFa: "نیکسو",             teamEn: "Nixo",                managerFa: "مهسا علمشاهی",         managerEn: "Mahsa Alamshahi",        username: "mahsa",     dept: "فروش" },
  { teamFa: "اینترپرایز A",      teamEn: "Enterprise A",        managerFa: "پیام درویشی",          managerEn: "Payam Darvishi",         username: "payam",     dept: "فروش" },
  { teamFa: "اینترپرایز B",      teamEn: "Enterprise B",        managerFa: "آدلین افراخته",        managerEn: "Adlin Afrakhteh",        username: "adlin",     dept: "فروش" },
  { teamFa: "ادزآن",             teamEn: "AdsOn",               managerFa: "محمدرضا خوش آمال",     managerEn: "Mohammadreza KhoshAmal", username: "khoshamal", dept: "فروش" },
  { teamFa: "پارتنرشیپ",         teamEn: "Partnership",         managerFa: "پویا تقوی",            managerEn: "Pouya Taghavi",          username: "pouya",     dept: "فروش" },
];

export const ADMIN_USER = { teamFa: "همه تیم‌ها", teamEn: "all", managerFa: "ادمین", managerEn: "Admin", username: "admin", dept: "" };

export interface SessionUser {
  username: string;
  managerFa: string;
  managerEn: string;
  teamFa: string;
  teamEn: string;
  isAdmin: boolean;
}
