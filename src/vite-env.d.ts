/// <reference types="vite/client" />
/// <reference types="vite-plugin-pwa/react" />

import 'react'

declare module 'react' {
  interface ButtonHTMLAttributes<T> extends HTMLAttributes<T> {
    commandfor?: string
    command?: string
  }
}
