
export interface Slide {
  title: string;
  content: string[];
  speakerNotes: string;
  translatedSpeakerNotes: string;
  imagePrompt: string;
  image?: string; // Base64 encoded image data
}
