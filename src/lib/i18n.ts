import i18n from "i18next";
import { initReactI18next } from "react-i18next";

import en from "@/locales/en.json";
import ja from "@/locales/ja.json";
import zhCN from "@/locales/zh-CN.json";

export const LOCALES = ["en", "zh-CN", "ja"] as const;

export type Locale = (typeof LOCALES)[number];

const STORAGE_KEY = "sheetlens.locale";

const resources = {
  en: { translation: en },
  "zh-CN": { translation: zhCN },
  ja: { translation: ja },
} as const;

function isLocale(value: unknown): value is Locale {
  return LOCALES.includes(value as Locale);
}

/** The language the user last picked, if they ever did. */
function storedLocale(): Locale | null {
  try {
    const value = localStorage.getItem(STORAGE_KEY);
    return isLocale(value) ? value : null;
  } catch {
    return null;
  }
}

/**
 * The first of the system's preferred languages that we ship, narrowed by the
 * language subtag alone — `zh-Hant-TW` and `zh-SG` both land on `zh-CN`, since
 * Simplified is the only Chinese we carry.
 */
function systemLocale(): Locale {
  const tags = navigator.languages?.length
    ? navigator.languages
    : [navigator.language];
  for (const tag of tags) {
    const [language] = tag.toLowerCase().split("-");
    if (language === "zh") return "zh-CN";
    if (language === "ja") return "ja";
    if (language === "en") return "en";
  }
  return "en";
}

export function initialLocale(): Locale {
  return storedLocale() ?? systemLocale();
}

// Registered before `init`, so it also relays the language chosen at startup.
i18n.on("languageChanged", (locale) => {
  document.documentElement.lang = locale;
});

void i18n.use(initReactI18next).init({
  resources,
  lng: initialLocale(),
  fallbackLng: "en",
  interpolation: { escapeValue: false },
});

export function setLocale(locale: Locale): void {
  try {
    localStorage.setItem(STORAGE_KEY, locale);
  } catch {
    // A webview that refuses storage still gets the language, just not kept.
  }
  void i18n.changeLanguage(locale);
}

export default i18n;
