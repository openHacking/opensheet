import type { SceneId } from './scenes.js';
export interface PlaygroundState {
  currentScene: SceneId;
  dirty: boolean;
  importedWorkbook: boolean;
  foundIndex: number;
}
export function createState(): PlaygroundState {
  return { currentScene: 'budget', dirty: false, importedWorkbook: false, foundIndex: -1 };
}
