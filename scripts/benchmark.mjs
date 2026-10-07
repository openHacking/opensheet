if (!process.argv.some((arg) => arg.startsWith('--sizes='))) process.argv.push('--sizes=10000');
await import('./browser-benchmark.mjs');
