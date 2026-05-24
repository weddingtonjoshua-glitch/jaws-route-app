"use client";
import dynamic from "next/dynamic";

const JAWSApp = dynamic(() => import("../components/JAWSApp"), { ssr: false });

export default function Page() {
  return <JAWSApp />;
}
