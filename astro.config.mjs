import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

export default defineConfig({
  site: 'https://technology-and-ai.seniorhousingsolutions.net',
  integrations: [sitemap()],
});
