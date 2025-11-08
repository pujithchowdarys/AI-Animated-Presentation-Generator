import { GoogleGenAI, Modality, Type } from "@google/genai";
import type { Slide } from "../types";

const API_KEY = process.env.API_KEY;
if (!API_KEY) {
  throw new Error("API_KEY environment variable is not set");
}

const ai = new GoogleGenAI({ apiKey: API_KEY });

const presentationOutputSchema = {
  type: Type.OBJECT,
  properties: {
    overallPresentationTitle: {
      type: Type.STRING,
      description: "A concise and engaging title for the entire presentation, derived from the topic. Max 100 characters.",
    },
    overallPresentationDescription: {
      type: Type.STRING,
      description: "A brief, compelling description of the presentation's scope and key takeaways, derived from the topic. Max 900 characters.",
    },
    slides: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          title: {
            type: Type.STRING,
            description: "The title of the slide. Should be concise and engaging.",
          },
          content: {
            type: Type.ARRAY,
            items: {
              type: Type.STRING,
            },
            description: "An array of strings, where each string is a bullet point for the slide content. Should contain 3-5 key points.",
          },
          speakerNotes: {
            type: Type.STRING,
            description: "Detailed speaker notes for the presenter in English. This will be shown on the presentation notes. Should be a full paragraph, elaborating on the bullet points.",
          },
          translatedSpeakerNotes: {
              type: Type.STRING,
              description: "The speaker notes from above, translated into the target voiceover language. If the target language is English, this should be the same as speakerNotes."
          },
          imagePrompt: {
            type: Type.STRING,
            description: "A detailed, visually descriptive prompt for an AI image generator to create a relevant, high-quality image for this slide. The prompt should describe a scene that visually represents the concept on the slide, like a metaphor or a diagram (e.g., 'A futuristic data pipeline showing glowing data streams flowing between servers'). This should create a background image for the slide content.",
          },
        },
        required: ["title", "content", "speakerNotes", "translatedSpeakerNotes", "imagePrompt"],
      },
    },
  },
  required: ["overallPresentationTitle", "overallPresentationDescription", "slides"],
};

export const generatePresentationContent = async (topic: string, voiceoverLanguage: string): Promise<{ overallPresentationTitle: string; overallPresentationDescription: string; slides: Slide[] }> => {
  let prompt = `Generate a 7-slide presentation about "${topic}".
First, provide a concise and engaging overall presentation title (max 100 characters) and a brief, compelling overall presentation description (max 900 characters) that summarizes the presentation's scope and key takeaways.
Then, for each of the 7 slides, provide: a title, 3-5 bullet points, detailed speaker notes in English, the speaker notes translated into ${voiceoverLanguage}, and a detailed image prompt.
All slide content (titles, bullet points) must be in English. The first slide should be a title slide, and the last should be a "Thank You" or "Q&A" slide.
The image prompt should describe a scene that visually represents the concept on the slide, like a metaphor or a diagram, to be used as a background image.`;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-2.5-pro",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: presentationOutputSchema,
      },
    });

    const jsonText = response.text.trim();
    const parsedResponse = JSON.parse(jsonText);
    
    if (!parsedResponse.overallPresentationTitle || !parsedResponse.overallPresentationDescription || !Array.isArray(parsedResponse.slides) || parsedResponse.slides.some((s: Slide) => !s.title || !s.content || !s.speakerNotes || !s.translatedSpeakerNotes || !s.imagePrompt)) {
        throw new Error("AI returned data in an unexpected format.");
    }

    return parsedResponse as { overallPresentationTitle: string; overallPresentationDescription: string; slides: Slide[] };
  } catch (error) {
    console.error("Error generating presentation content:", error);
    throw new Error("Failed to generate presentation content from AI.");
  }
};

export const generateImage = async (prompt: string): Promise<string> => {
  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash-image',
      contents: {
        parts: [{ text: prompt }],
      },
      config: {
        responseModalities: [Modality.IMAGE],
      },
    });
    for (const part of response.candidates[0].content.parts) {
      if (part.inlineData) {
        return part.inlineData.data;
      }
    }
    throw new Error("No image data received from API.");
  } catch (error) {
    console.error("Error generating image:", error);
    throw new Error("Failed to generate image.");
  }
};

export const generateVoiceover = async (text: string, voice: string): Promise<string> => {
  try {
    const response = await ai.models.generateContent({
      model: "gemini-2.5-flash-preview-tts",
      contents: [{ parts: [{ text: `Say with a professional and clear tone: ${text}` }] }],
      config: {
        responseModalities: [Modality.AUDIO],
        speechConfig: {
          voiceConfig: {
            prebuiltVoiceConfig: { voiceName: voice },
          },
        },
      }, // Added the missing closing curly brace for the config object
    });
    
    const audioData = response.candidates?.[0]?.content?.parts?.[0]?.inlineData?.data;
    if (!audioData) {
      throw new Error("No audio data received from API.");
    }
    return audioData;
  } catch (error) {
    console.error("Error generating voiceover:", error);
    throw new Error("Failed to generate voiceover.");
  }
};