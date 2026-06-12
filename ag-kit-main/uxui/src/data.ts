export type VulnSeverity = "critical" | "high" | "moderate" | "none";
export type PackageCategory =
  | "core"
  | "framework"
  | "styling"
  | "tooling"
  | "testing"
  | "utility";

export interface PackageNode {
  id: string;
  name: string;
  version: string;
  description: string;
  category: PackageCategory;
  x: number;
  y: number;
  dependencies: string[];
  vulnerability: VulnSeverity;
  vulnCount?: number;
  outdated?: boolean;
  license: string;
  size?: string;
  colorDark?: string;
  colorLight?: string;
}

export interface Dependency {
  from: string;
  to: string;
}

export const CATEGORY_META: Record<
  PackageCategory,
  { label: string; icon: string }
> = {
  core: { label: "Core", icon: "atom" },
  framework: { label: "Frameworks", icon: "triangle" },
  styling: { label: "CSS", icon: "brush" },
  tooling: { label: "Dev Tools", icon: "container" },
  testing: { label: "Testing", icon: "flask" },
  utility: { label: "Libraries", icon: "plug" },
};

export const PACKAGES: PackageNode[] = [
  // Row 1 — Core
  {
    id: "react",
    name: "react",
    version: "19.2.4",
    description:
      "A JavaScript library for building user interfaces using declarative components",
    category: "core",
    x: 30,
    y: 14,
    dependencies: [],
    vulnerability: "none",
    license: "MIT",
    size: "6.4 kB",
    colorDark: "#61dafb",
    colorLight: "#0c7bb3",
  },
  {
    id: "react-dom",
    name: "react-dom",
    version: "19.2.4",
    description:
      "React package for working with the DOM — renders components to the browser",
    category: "core",
    x: 70,
    y: 14,
    dependencies: ["react"],
    vulnerability: "none",
    license: "MIT",
    size: "137 kB",
    colorDark: "#61dafb",
    colorLight: "#0c7bb3",
  },

  // Row 2 — Framework & Styling
  {
    id: "nextjs",
    name: "next",
    version: "15.3.1",
    description:
      "The React framework for production — SSR, routing, and optimized builds",
    category: "framework",
    x: 56,
    y: 30,
    dependencies: ["react", "react-dom"],
    vulnerability: "none",
    license: "MIT",
    size: "5.2 MB",
  },
  {
    id: "tailwindcss",
    name: "tailwindcss",
    version: "4.2.1",
    description:
      "A utility-first CSS framework for rapidly building custom designs",
    category: "styling",
    x: 88,
    y: 60,
    dependencies: [],
    vulnerability: "none",
    license: "MIT",
    size: "490 kB",
  },
  {
    id: "framer-motion",
    name: "framer-motion",
    version: "12.38.0",
    description:
      "Production-ready motion library for React with spring animations",
    category: "utility",
    x: 82,
    y: 31,
    dependencies: ["react", "react-dom"],
    vulnerability: "none",
    license: "MIT",
    size: "140 kB",
  },

  // Row 3 — Utilities & Data
  {
    id: "axios",
    name: "axios",
    version: "1.7.2",
    description: "Promise-based HTTP client for the browser and Node.js",
    category: "utility",
    x: 14,
    y: 82,
    dependencies: [],
    vulnerability: "moderate",
    vulnCount: 1,
    license: "MIT",
    size: "29 kB",
  },
  {
    id: "zustand",
    name: "zustand",
    version: "5.0.5",
    description:
      "Small, fast, and scalable bearbones state management for React",
    category: "utility",
    x: 22,
    y: 44,
    dependencies: ["react"],
    vulnerability: "none",
    license: "MIT",
    size: "3.1 kB",
  },
  {
    id: "zod",
    name: "zod",
    version: "3.24.4",
    description:
      "TypeScript-first schema validation with static type inference",
    category: "utility",
    x: 74,
    y: 47,
    dependencies: [],
    vulnerability: "none",
    license: "MIT",
    size: "13 kB",
  },
  {
    id: "date-fns",
    name: "date-fns",
    version: "2.30.0",
    description:
      "Modern JavaScript date utility library — modular & tree-shakeable",
    category: "utility",
    x: 92,
    y: 43,
    dependencies: [],
    vulnerability: "none",
    outdated: true,
    license: "MIT",
    size: "75 kB",
  },

  // Row 4 — Tooling
  {
    id: "typescript",
    name: "typescript",
    version: "5.9.3",
    description:
      "Typed superset of JavaScript that compiles to plain JavaScript",
    category: "tooling",
    x: 12,
    y: 56,
    dependencies: [],
    vulnerability: "none",
    license: "Apache-2.0",
    size: "22 MB",
    colorDark: "#3178c6",
    colorLight: "#1b5a9e",
  },
  {
    id: "vite",
    name: "vite",
    version: "8.0.0",
    description:
      "Next-generation frontend tooling — blazing fast HMR and build",
    category: "tooling",
    x: 58,
    y: 68,
    dependencies: [],
    vulnerability: "none",
    license: "MIT",
    size: "4.8 MB",
  },
  {
    id: "eslint",
    name: "eslint",
    version: "10.0.3",
    description: "Pluggable linting utility for JavaScript and TypeScript code",
    category: "tooling",
    x: 10,
    y: 32,
    dependencies: [],
    vulnerability: "none",
    license: "MIT",
    size: "3.1 MB",
  },

  // Row 5 — Build plugins
  {
    id: "vite-plugin-react",
    name: "@vitejs/plugin-react",
    version: "6.0.1",
    description:
      "Official Vite plugin for React — Fast Refresh and JSX transform",
    category: "tooling",
    x: 32,
    y: 64,
    dependencies: ["vite", "react"],
    vulnerability: "none",
    license: "MIT",
    size: "210 kB",
  },
  {
    id: "tailwind-vite",
    name: "@tailwindcss/vite",
    version: "4.2.1",
    description: "First-party Tailwind CSS integration for Vite projects",
    category: "styling",
    x: 86,
    y: 82,
    dependencies: ["vite", "tailwindcss"],
    vulnerability: "none",
    license: "MIT",
    size: "45 kB",
  },

  // Row 6 — Testing
  {
    id: "vitest",
    name: "vitest",
    version: "3.2.1",
    description: "Blazing-fast unit testing framework powered by Vite",
    category: "testing",
    x: 35,
    y: 82,
    dependencies: ["vite"],
    vulnerability: "none",
    license: "MIT",
    size: "8.3 MB",
  },
  {
    id: "testing-library",
    name: "@testing-library/react",
    version: "16.3.0",
    description: "Simple and complete React DOM testing utilities",
    category: "testing",
    x: 58,
    y: 52,
    dependencies: ["react", "react-dom"],
    vulnerability: "none",
    license: "MIT",
    size: "120 kB",
  },
  {
    id: "playwright",
    name: "@playwright/test",
    version: "1.52.0",
    description: "End-to-end testing framework for modern web apps",
    category: "testing",
    x: 62,
    y: 82,
    dependencies: [],
    vulnerability: "high",
    vulnCount: 2,
    license: "Apache-2.0",
    size: "95 MB",
  },
];

export const DEPENDENCIES: Dependency[] = [
  { from: "react-dom", to: "react" },
  { from: "nextjs", to: "react" },
  { from: "nextjs", to: "react-dom" },
  { from: "framer-motion", to: "react" },
  { from: "framer-motion", to: "react-dom" },
  { from: "zustand", to: "react" },
  { from: "vite-plugin-react", to: "vite" },
  { from: "vite-plugin-react", to: "react" },
  { from: "tailwind-vite", to: "vite" },
  { from: "tailwind-vite", to: "tailwindcss" },
  { from: "vitest", to: "vite" },
  { from: "testing-library", to: "react" },
  { from: "testing-library", to: "react-dom" },
];
