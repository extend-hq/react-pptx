import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { ThemeProvider } from 'next-themes';
import '@extend-ai/react-pptx/styles.css';
import { TooltipProvider } from '@/components/ui/tooltip';
import { App } from './App';
import './index.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ThemeProvider
      attribute="class"
      defaultTheme="light"
      disableTransitionOnChange
      enableSystem={false}
    >
      <TooltipProvider delay={400}>
        <App />
      </TooltipProvider>
    </ThemeProvider>
  </StrictMode>,
);
