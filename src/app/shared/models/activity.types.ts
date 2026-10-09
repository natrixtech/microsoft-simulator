export type AppKind = 'word' | 'excel' | 'powerpoint';

export type ThemePreference = 'LIGHT' | 'DARK';

export interface BaseGoal {
  type: string;
  [key: string]: unknown;
}

export interface TaskGoal extends BaseGoal {
  type: string;
}

export interface LessonTask {
  title: string;
  explanation: string;
  instructions: string[];
  hint?: string;
  goals: TaskGoal[];
  showMeHow?: boolean;
  startingDocument?: string | null;
  startingSheet?: any | null;
  startingDeck?: any | null;
  pageSetup?: {
    orientation?: 'portrait' | 'landscape';
    paperSize?: 'Letter' | 'A4' | 'Legal';
    margins?: { top: number; bottom: number; left: number; right: number };
  };
}

export interface ActivityPayload {
  assetBaseUrl: string;
  theme: ThemePreference;
  mode?: 'lesson' | 'assessment';
  title?: string;
  instruction?: string;
  startingDocument?: string;
  goals?: TaskGoal[];
  tasks?: LessonTask[];
  showMeHow?: boolean;
  pageSetup?: {
    orientation?: 'portrait' | 'landscape';
    paperSize?: 'Letter' | 'A4' | 'Legal';
    margins?: { top: number; bottom: number; left: number; right: number };
  };
  images?: any;
  comments?: any;
  variables?: Record<string, string>;
  [key: string]: unknown;
}

export interface PluginEventMessage {
  type: 'PLUGIN_EVENT';
  name: 'pluginReady' | 'activityCompleted' | 'pluginError';
  payload?: {
    completed: boolean;
    tasksPassed?: number;
    tasksTotal?: number;
  };
  message?: string;
}

export type IncomingPluginMessage =
  | { type: 'PLUGIN_DATA_IN'; payload: ActivityPayload }
  | { type: 'PAUSE'; payload: { pause: boolean } }
  | { type: 'FORCE_FINISH' }
  | { type: 'PREFERENCE_CHANGED'; payload: { type: 'theme'; value: ThemePreference } };
