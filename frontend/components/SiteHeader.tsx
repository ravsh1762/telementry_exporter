"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { NAV_ITEMS } from "@/lib/tools";
import { ThemeToggle } from "./ThemeToggle";

export function SiteHeader({ subtitle }: { subtitle: string }) {
  const pathname = usePathname();

  return (
    <header className="border-b border-slate-200 bg-white/90 backdrop-blur-md sticky top-0 z-50">
      <div className="max-w-[1600px] mx-auto px-6 min-h-14 py-2 flex flex-wrap lg:flex-nowrap items-center gap-x-6 gap-y-2">
        {/* brand */}
        <div className="flex items-center gap-3 mr-auto lg:mr-0">
          <div className="w-7 h-7 rounded-lg bg-blue-600 flex items-center justify-center shadow-sm">
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
              <path
                d="M7 1L12.5 4.25V10.75L7 14L1.5 10.75V4.25L7 1Z"
                stroke="white"
                strokeWidth="1.3"
              />
              <path d="M7 4L9.5 5.5V8.5L7 10L4.5 8.5V5.5L7 4Z" fill="white" />
            </svg>
          </div>
          <div className="flex items-baseline gap-2">
            <span className="font-bold text-slate-800 text-sm tracking-tight">
              Telemetry Exporter
            </span>
            <span className="text-slate-300 text-xs hidden 2xl:inline">·</span>
            <span className="text-slate-400 text-xs hidden 2xl:inline">
              {subtitle}
            </span>
          </div>
        </div>

        {/* nav – drops to its own row on small screens, sits right on desktop */}
        <nav className="order-3 lg:order-2 w-full lg:w-auto lg:ml-auto overflow-x-auto">
          <ul className="flex items-center gap-1 min-w-max">
            {NAV_ITEMS.map((item) => {
              const active = pathname === item.href;
              return (
                <li key={item.href}>
                  <Link
                    href={item.href}
                    className={`px-3.5 py-1.5 text-xs font-semibold rounded-full whitespace-nowrap transition-all border ${
                      active
                        ? "bg-blue-600 text-white border-blue-500 shadow-sm"
                        : "text-slate-500 border-transparent hover:text-slate-800 hover:border-slate-200 hover:bg-slate-100"
                    }`}
                  >
                    {item.label}
                  </Link>
                </li>
              );
            })}
          </ul>
        </nav>

        {/* status + theme switch */}
        <div className="order-2 lg:order-3 flex items-center gap-4">
          <div className="hidden 2xl:flex items-center gap-1.5 text-xs text-slate-500">
            <div className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
            System Online
          </div>
          <ThemeToggle />
        </div>
      </div>
    </header>
  );
}
