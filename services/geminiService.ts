
import { GoogleGenAI, Modality, Type } from "@google/genai";
import type { Slide } from "../types";

const API_KEY = process.env.API_KEY;
if (!API_KEY) {
  throw new Error("API_KEY environment variable is not set");
}

const ai = new GoogleGenAI({ apiKey: API_KEY });

const presentationSchema = {
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
        description: "Detailed speaker notes for the presenter. This will be converted to audio. Should be a full paragraph, elaborating on the bullet points.",
      },
    },
    required: ["title", "content", "speakerNotes"],
  },
};

export const generatePresentationContent = async (topic: string): Promise<Slide[]> => {
  const prompt = `Create a 7-slide presentation about "${topic}". The first slide should be a title slide, and the last should be a "Thank You" or "Q&A" slide. For each slide, provide a title, 3-5 bullet points for the main content, and detailed speaker notes. The content should be informative and well-structured.`;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-2.5-pro",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: presentationSchema,
      },
    });

    const jsonText = response.text.trim();
    const presentationData = JSON.parse(jsonText);
    
    // Basic validation
    if (!Array.isArray(presentationData) || presentationData.some(s => !s.title || !s.content || !s.speakerNotes)) {
        throw new Error("AI returned data in an unexpected format.");
    }

    return presentationData as Slide[];
  } catch (error) {
    console.error("Error generating presentation content:", error);
    throw new Error("Failed to generate presentation content from AI.");
  }
};

export const generateVoiceover = async (text: string): Promise<string> => {
  try {
    const response = await ai.models.generateContent({
      model: "gemini-2.5-flash-preview-tts",
      contents: [{ parts: [{ text: text }] }],
      config: {
        responseModalities: [Modality.AUDIO],
        speechConfig: {
          voiceConfig: {
            prebuiltVoiceConfig: { voiceName: 'Kore' },
          },
        },
      },
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
