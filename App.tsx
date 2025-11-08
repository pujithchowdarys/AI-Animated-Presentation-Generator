import React, { useState, useCallback, useRef, useEffect } from 'react';
import PptxGenJS from 'pptxgenjs';
import { generatePresentationContent, generateVoiceover, generateImage, getGenerativeModel } from './services/geminiService';
import type { Slide } from './types';
import { SlidePreview } from './components/SlidePreview';
import { Loader } from './components/Loader';
import { DownloadIcon, PresentationIcon, PlayIcon, PauseIcon, AudioIcon, VideoIcon } from './components/Icons';
import { audioBufferToWavBlob, decode, decodeAudioData } from './utils/audioUtils';

// Define a type that can be either a standard or an offscreen canvas rendering context
type GenericCanvasRenderingContext2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

const voiceoverLanguages = [
  { code: 'English', name: 'English' },
  { code: 'Hindi', name: 'Hindi' },
  { code: 'Tamil', name: 'Tamil' },
  { code: 'Telugu', name: 'Telugu' },
];

const voices = [
  { id: 'Kore', name: 'Kore (Female)' },
  { id: 'Puck', name: 'Puck (Male)' },
  { id: 'Charon', name: 'Charon (Male)' },
  { id: 'Fenrir', name: 'Fenrir (Male)' },
  { id: 'Zephyr', name: 'Zephyr (Female)' },
];

const LOGICAL_CANVAS_WIDTH = 1280; // Standard 16:9 width for drawing logic
const LOGICAL_CANVAS_HEIGHT = 720; // Standard 16:9 height for drawing logic
const VIDEO_OUTPUT_WIDTH = 1920; // Target 1080p width
const VIDEO_OUTPUT_HEIGHT = 1080; // Target 1080p height
const SCALE_FACTOR_X = VIDEO_OUTPUT_WIDTH / LOGICAL_CANVAS_WIDTH;
const SCALE_FACTOR_Y = VIDEO_OUTPUT_HEIGHT / LOGICAL_CANVAS_HEIGHT;

const FONT_TITLE = 'Inter-Bold';
const FONT_BODY = 'Inter';

// Helper to draw text with wrapping
const drawText = (ctx: GenericCanvasRenderingContext2D, text: string, x: number, y: number, maxWidth: number, lineHeight: number, color: string, fontSize: number, bold: boolean, align: CanvasTextAlign = 'center') => {
  ctx.fillStyle = color;
  ctx.font = `${bold ? 'bold ' : ''}${fontSize}px ${bold ? FONT_TITLE : FONT_BODY}, sans-serif`;
  ctx.textAlign = align;
  ctx.textBaseline = 'top'; // Changed to top for consistent line height calculation

  const words = text.split(' ');
  let line = '';
  let currentY = y;
  
  if (align === 'center') x += maxWidth / 2; // Adjust x for center alignment

  for (let n = 0; n < words.length; n++) {
    const testLine = line + words[n] + ' ';
    const metrics = ctx.measureText(testLine);
    const testWidth = metrics.width;
    if (testWidth > maxWidth && n > 0) {
      ctx.fillText(line.trim(), x, currentY);
      line = words[n] + ' ';
      currentY += lineHeight;
    } else {
      line = testLine;
    }
  }
  ctx.fillText(line.trim(), x, currentY);
  return currentY; // Return the final Y position
};

// Helper for drawing bullet points
const drawBulletPoints = (ctx: GenericCanvasRenderingContext2D, points: string[], x: number, y: number, maxWidth: number, lineHeight: number, bulletSize: number, textColor: string, bulletColor: string, fontSize: number) => {
  let currentY = y;
  points.forEach(point => {
    ctx.fillStyle = bulletColor;
    ctx.font = `${bulletSize}px ${FONT_BODY}, sans-serif`;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
    ctx.fillText('•', x, currentY); // Bullet point

    const textX = x + bulletSize * 1.5;
    const textMaxWidth = maxWidth - bulletSize * 1.5;
    const finalY = drawText(ctx, point, textX, currentY, textMaxWidth, lineHeight, textColor, fontSize, false, 'left');
    currentY = finalY + lineHeight; // Move to next line for next bullet
  });
};


const App: React.FC = () => {
  const [topic, setTopic] = useState<string>('');
  const [generatedPresentationTitle, setGeneratedPresentationTitle] = useState<string>(''); // New state for AI-generated presentation title
  const [generatedPresentationDescription, setGeneratedPresentationDescription] = useState<string>(''); // New state for AI-generated presentation description
  const [voiceoverLanguage, setVoiceoverLanguage] = useState<string>(voiceoverLanguages[0].code);
  const [voice, setVoice] = useState<string>(voices[0].id);
  const [slides, setSlides] = useState<Slide[]>([]);
  const [currentSlide, setCurrentSlide] = useState<number>(0);
  const [isLoading, setIsLoading] = useState<boolean>(true); // Start as true for API key check
  const [loadingMessage, setLoadingMessage] = useState<string>('Initializing application...');
  const [error, setError] = useState<string | null>(null);
  const [audioUrl, setAudioUrl] = useState<string | null>(null);
  const [isPlaying, setIsPlaying] = useState<boolean>(false);
  const [isGeneratingVideo, setIsGeneratingVideo] = useState<boolean>(false);
  const [videoDownloadUrl, setVideoDownloadUrl] = useState<string | null>(null);
  const [isServiceInitialized, setIsServiceInitialized] = useState<boolean>(false); // New state for API key check

  const audioRef = useRef<HTMLAudioElement>(null);
  const audioBuffersRef = useRef<AudioBuffer[]>([]); // To store individual audio buffers for video generation

  // Effect to check API key on component mount
  useEffect(() => {
    const initializeService = async () => {
      try {
        setLoadingMessage('Initializing application...');
        await getGenerativeModel(); // Attempt to get the GenAI model instance
        setIsServiceInitialized(true);
        setError(null);
      } catch (e: any) {
        setIsServiceInitialized(false);
        setError(e.message);
        console.error("API Key Initialization Error:", e);
      } finally {
        setIsLoading(false);
        setLoadingMessage('');
      }
    };
    initializeService();
  }, []);

  const drawSlideToCanvas = useCallback(async (ctx: GenericCanvasRenderingContext2D, slide: Slide) => {
    ctx.clearRect(0, 0, LOGICAL_CANVAS_WIDTH, LOGICAL_CANVAS_HEIGHT);

    // Background
    if (slide.image) {
      const img = new Image();
      img.src = `data:image/png;base64,${slide.image}`;
      await new Promise(resolve => img.onload = resolve);
      ctx.drawImage(img, 0, 0, LOGICAL_CANVAS_WIDTH, LOGICAL_CANVAS_HEIGHT);
      
      // Overlay
      ctx.fillStyle = 'rgba(0, 0, 0, 0.6)'; // Semi-transparent black
      ctx.fillRect(0, 0, LOGICAL_CANVAS_WIDTH, LOGICAL_CANVAS_HEIGHT);
    } else {
      ctx.fillStyle = '#1A202C'; // bg-slate-900 equivalent
      // Fix: Corrected typo 'LOGICAL_CANAS_HEIGHT' to 'LOGICAL_CANVAS_HEIGHT'
      ctx.fillRect(0, 0, LOGICAL_CANVAS_WIDTH, LOGICAL_CANVAS_HEIGHT);
    }

    // Title
    drawText(ctx, slide.title, 50, 50, LOGICAL_CANVAS_WIDTH - 100, 60, '#38BDF8', 44, true, 'center');

    // Content
    const contentStartX = 100;
    const contentStartY = 200;
    const contentMaxWidth = LOGICAL_CANVAS_WIDTH - 200;
    const contentLineHeight = 36;
    const bulletSize = 24;
    drawBulletPoints(ctx, slide.content, contentStartX, contentStartY, contentMaxWidth, contentLineHeight, bulletSize, '#E2E8F0', '#E2E8F0', 24);

  }, []);

  const handleGenerate = async () => {
    if (!topic.trim()) {
      setError('Please enter a topic.');
      return;
    }
    setIsLoading(true);
    setLoadingMessage('Warming up the AI engine...');
    setError(null);
    setSlides([]);
    setGeneratedPresentationTitle('');
    setGeneratedPresentationDescription('');
    setAudioUrl(null);
    setVideoDownloadUrl(null); // Clear previous video URL
    setCurrentSlide(0);
    audioBuffersRef.current = []; // Clear previous audio buffers

    try {
      setLoadingMessage('Generating presentation content and details...');
      // Pass new title and description to the service
      const { overallPresentationTitle, overallPresentationDescription, slides: generatedSlides } = await generatePresentationContent(topic, voiceoverLanguage);
      
      setGeneratedPresentationTitle(overallPresentationTitle);
      setGeneratedPresentationDescription(overallPresentationDescription);
      setSlides(generatedSlides);

      const slidesWithImages: Slide[] = [...generatedSlides];
      for (let i = 0; i < generatedSlides.length; i++) {
        setLoadingMessage(`Generating image for slide ${i + 1}/${generatedSlides.length}...`);
        const imageBase64 = await generateImage(generatedSlides[i].imagePrompt);
        slidesWithImages[i] = { ...generatedSlides[i], image: imageBase64 };
        setSlides([...slidesWithImages]);
      }

      setLoadingMessage('Crafting voiceover for each slide...');
      const audioContext = new (window.AudioContext || (window as any).webkitAudioContext)();
      let totalLength = 0;

      for (let i = 0; i < slidesWithImages.length; i++) {
        setLoadingMessage(`Generating voice for slide ${i + 1}/${slidesWithImages.length}...`);
        const slide = slidesWithImages[i];
        if (slide.translatedSpeakerNotes) {
          const audioBase64 = await generateVoiceover(slide.translatedSpeakerNotes, voice);
          const audioBytes = decode(audioBase64);
          const audioBuffer = await decodeAudioData(audioBytes, audioContext, 24000, 1);
          audioBuffersRef.current.push(audioBuffer); // Store individual buffer
          totalLength += audioBuffer.length;
        }
      }
      
      setLoadingMessage('Stitching audio clips together for combined playback...');
      const mergedBuffer = audioContext.createBuffer(1, totalLength, 24000);
      const channelData = mergedBuffer.getChannelData(0);
      let offset = 0;
      for (const buffer of audioBuffersRef.current) { // Use stored buffers for merged audio
        channelData.set(buffer.getChannelData(0), offset);
        offset += buffer.length;
      }

      const wavBlob = audioBufferToWavBlob(mergedBuffer);
      const url = URL.createObjectURL(wavBlob);
      setAudioUrl(url);

    } catch (e: any) {
      setError(`An error occurred: ${e.message}`);
      console.error(e);
    } finally {
      setIsLoading(false);
      setLoadingMessage('');
    }
  };

  const handleDownloadPpt = useCallback(() => {
    if (slides.length === 0) return;

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';

    // Add presentation title and description to the PPTX if they exist
    if (generatedPresentationTitle || generatedPresentationDescription) {
        const titleSlide = pptx.addSlide();
        titleSlide.background = { color: '1A202C' }; // Dark background
        if (generatedPresentationTitle) {
            titleSlide.addText(generatedPresentationTitle, {
                x: 0.5, y: 1.5, w: '90%', h: 1,
                fontSize: 48, bold: true, color: '38BDF8', align: 'center',
                objectName: 'generated_presentation_title',
            });
        }
        if (generatedPresentationDescription) {
            titleSlide.addText(generatedPresentationDescription, {
                x: 0.5, y: 3, w: '90%', h: 1.5,
                fontSize: 24, color: 'E2E8F0', align: 'center',
                objectName: 'generated_presentation_description',
            });
        }
    }


    slides.forEach((slide) => {
      const pptSlide = pptx.addSlide();
      
      if (slide.image) {
        pptSlide.background = { data: `data:image/png;base64,${slide.image}` };
        // Add a semi-transparent overlay for text readability
        pptSlide.addShape(PptxGenJS.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '000000', transparency: 60 } });
        
        pptSlide.addText(slide.title, { 
          x: 0.5, y: 0.25, w: '90%', h: 1, 
          fontSize: 44, bold: true, color: 'FFFFFF', align: 'center',
          objectName: 'title',
        });

        const contentPoints = slide.content.map(point => ({ text: point, options: { bullet: true, color: 'E2E8F0' } }));
        pptSlide.addText(contentPoints, {
          x: 1, y: 1.5, w: '80%', h: 3.5,
          fontSize: 24, color: 'FFFFFF', lineSpacing: 36,
          objectName: 'content',
        });

      } else {
        pptSlide.background = { color: '1A202C' };
        pptSlide.addText(slide.title, { 
          x: 0.5, y: 0.25, w: '90%', h: 1, 
          fontSize: 36, bold: true, color: '38BDF8', align: 'center',
          objectName: 'title',
        });

        const contentPoints = slide.content.map(point => ({ text: point, options: { bullet: true, color: 'E2E8F0' } }));
        pptSlide.addText(contentPoints, {
          x: 1, y: 1.5, w: '80%', h: 3.5,
          fontSize: 20, lineSpacing: 30,
          objectName: 'content',
        });
      }
      pptSlide.addNotes(slide.speakerNotes);
    });

    pptx.writeFile({ fileName: `${topic.replace(/\s+/g, '_')}_presentation.pptx` });
  }, [slides, topic, generatedPresentationTitle, generatedPresentationDescription]);

  const handleDownloadVideo = useCallback(async () => {
    if (slides.length === 0 || audioBuffersRef.current.length === 0) {
      setError('Please generate a presentation first.');
      return;
    }

    setIsGeneratingVideo(true);
    setLoadingMessage('Preparing for 1080p video generation...');
    setError(null);
    setVideoDownloadUrl(null);

    // Declared here to be accessible in cleanupCanvas
    let visibleCanvas: HTMLCanvasElement | null = null; 

    const cleanupCanvas = () => {
      if (visibleCanvas && document.body.contains(visibleCanvas)) {
        document.body.removeChild(visibleCanvas);
      }
    };

    try {
      // Ensure fonts are loaded before drawing on canvas
      await document.fonts.ready;
      setLoadingMessage('Fonts loaded. Starting 1080p video rendering...');

      // Always create a visible HTMLCanvasElement for `captureStream`
      visibleCanvas = document.createElement('canvas');
      visibleCanvas.width = VIDEO_OUTPUT_WIDTH; // Use higher resolution
      visibleCanvas.height = VIDEO_OUTPUT_HEIGHT; // Use higher resolution
      visibleCanvas.style.display = 'none'; // Keep it hidden
      document.body.appendChild(visibleCanvas);
      const visibleCtx = visibleCanvas.getContext('2d', { willReadFrequently: true });
      if (!visibleCtx) {
        throw new Error("Failed to get 2D context for visible canvas.");
      }
      visibleCtx.scale(SCALE_FACTOR_X, SCALE_FACTOR_Y); // Apply scaling to visible context

      // Determine the actual drawing canvas and its context
      let drawingCanvas: HTMLCanvasElement | OffscreenCanvas;
      let drawingCtx: GenericCanvasRenderingContext2D;

      if (typeof OffscreenCanvas !== 'undefined') {
        drawingCanvas = new OffscreenCanvas(VIDEO_OUTPUT_WIDTH, VIDEO_OUTPUT_HEIGHT); // Use higher resolution
        drawingCtx = drawingCanvas.getContext('2d', { willReadFrequently: true }) as OffscreenCanvasRenderingContext2D;
        setLoadingMessage('Using OffscreenCanvas for rendering...');
      } else {
        // If OffscreenCanvas is not available, draw directly to the visible canvas
        drawingCanvas = visibleCanvas;
        drawingCtx = visibleCtx as CanvasRenderingContext2D;
        setLoadingMessage('Using standard Canvas for rendering...');
      }

      if (!drawingCtx) {
        throw new Error("Failed to get 2D context for drawing.");
      }
      
      // Apply scaling to the drawing context if it's separate from the visible one
      if (drawingCanvas !== visibleCanvas) {
        drawingCtx.scale(SCALE_FACTOR_X, SCALE_FACTOR_Y);
      }


      // Robust check for captureStream availability on the *visible* canvas
      if (typeof visibleCanvas.captureStream !== 'function') {
        throw new Error("Your browser/environment does not support canvas.captureStream(), which is required for video generation.");
      }

      const canvasVideoStream = visibleCanvas.captureStream(30); // 30 FPS for video
      
      const audioContext = new (window.AudioContext || (window as any).webkitAudioContext)();
      const audioDestination = audioContext.createMediaStreamDestination();
      
      // Create a new MediaStream to combine video from canvas and audio from AudioContext
      const combinedStream = new MediaStream();
      canvasVideoStream.getVideoTracks().forEach(track => combinedStream.addTrack(track));
      audioDestination.stream.getAudioTracks().forEach(track => combinedStream.addTrack(track));


      const mediaRecorder = new MediaRecorder(combinedStream, { 
        mimeType: 'video/webm; codecs="vp8, opus"', // Specify codecs for quality and compatibility
        videoBitsPerSecond: 8_000_000, // 8 Mbps for 1080p
        audioBitsPerSecond: 128_000, // 128 Kbps
      });
      const recordedChunks: Blob[] = [];

      mediaRecorder.ondataavailable = (event) => {
        if (event.data.size > 0) {
          recordedChunks.push(event.data);
        }
      };

      const videoGenerationPromise = new Promise<string>((resolve) => {
        mediaRecorder.onstop = () => {
          const blob = new Blob(recordedChunks, { type: 'video/webm' });
          const url = URL.createObjectURL(blob);
          cleanupCanvas();
          resolve(url);
        };
      });

      mediaRecorder.start();

      for (let i = 0; i < slides.length; i++) {
        setLoadingMessage(`Rendering slide ${i + 1}/${slides.length} and playing voiceover for 1080p video...`);
        // Draw to the offscreen/drawing canvas
        await drawSlideToCanvas(drawingCtx, slides[i]);

        if (drawingCanvas instanceof OffscreenCanvas) {
          // Transfer the rendered image from offscreen to the visible canvas for capture
          // Draw the full-resolution bitmap onto the scaled visible canvas, filling it.
          visibleCtx.drawImage(drawingCanvas.transferToImageBitmap(), 0, 0, LOGICAL_CANVAS_WIDTH, LOGICAL_CANVAS_HEIGHT);
        }
        // If drawingCanvas is visibleCanvas, the drawing is already directly on it.

        const audioBuffer = audioBuffersRef.current[i];
        if (audioBuffer) {
          const source = audioContext.createBufferSource();
          source.buffer = audioBuffer;
          source.connect(audioDestination);
          source.start();

          // Wait for audio to finish playing
          await new Promise<void>((resolve) => {
            source.onended = () => {
              source.disconnect();
              resolve();
            };
          });
        }
      }

      mediaRecorder.stop();
      const finalVideoUrl = await videoGenerationPromise;
      setVideoDownloadUrl(finalVideoUrl);
      setLoadingMessage('1080p Video generation complete!');

    } catch (e: any) {
      setError(`Error generating video: ${e.message}`);
      console.error('Video generation error:', e);
      cleanupCanvas(); // Ensure cleanup on error
    } finally {
      setIsGeneratingVideo(false);
    }
  }, [slides, audioBuffersRef, topic, drawSlideToCanvas]);


  const toggleAudio = () => {
    if (!audioRef.current) return;
    if (isPlaying) {
      audioRef.current.pause();
    } else {
      audioRef.current.play();
    }
    setIsPlaying(!isPlaying);
  };
  
  const hasResults = slides.length > 0;
  // Use isServiceInitialized to gate the main app content
  const isBusy = isLoading || isGeneratingVideo || !isServiceInitialized;

  return (
    <div className="min-h-screen bg-slate-900 flex flex-col items-center p-4 sm:p-6 md:p-8">
      <header className="w-full max-w-5xl text-center mb-8">
        <div className="flex justify-center items-center gap-4 mb-2">
            <PresentationIcon className="w-10 h-10 text-sky-400" />
            <h1 className="text-3xl sm:text-4xl font-bold tracking-tight text-white">Pujiverse Presentation Generator</h1>
        </div>
        <p className="text-slate-400 text-lg">Turn any topic into a downloadable presentation with voiceover and video in seconds.</p>
      </header>
      
      <main className="w-full max-w-5xl flex-1 flex flex-col">
        {isLoading && !error && <Loader message={loadingMessage} />}

        {error && (
          <div className="bg-red-900/50 rounded-lg p-6 shadow-2xl border border-red-700 backdrop-blur-sm text-center my-8 animate-fade-in-up">
            <p className="text-red-300 text-lg mb-4">
              A critical error occurred: <strong className="text-red-100">{error}</strong>
            </p>
            {error.includes("API_KEY environment variable is not set") && (
              <p className="text-red-200 text-md">
                Please ensure your `API_KEY` environment variable is correctly configured.
                If deploying on Vercel, navigate to your project settings, then "Environment Variables", and add `API_KEY` with your Google Gemini API key.
              </p>
            )}
          </div>
        )}

        {isServiceInitialized && !error && (
          <>
            <div className="bg-slate-800/50 rounded-lg p-6 shadow-2xl border border-slate-700 backdrop-blur-sm">
                <div className="flex flex-col gap-4">
                  <div className="flex flex-col sm:flex-row gap-4">
                      <input
                          type="text"
                          value={topic}
                          onChange={(e) => setTopic(e.target.value)}
                          placeholder="e.g., The Future of Renewable Energy (main topic)"
                          className="flex-grow bg-slate-700 text-white placeholder-slate-400 rounded-md px-4 py-3 border border-slate-600 focus:ring-2 focus:ring-sky-500 focus:outline-none transition"
                          disabled={isLoading}
                      />
                      <button
                          onClick={handleGenerate}
                          disabled={isLoading || !topic.trim()}
                          className="flex justify-center items-center gap-2 bg-sky-500 hover:bg-sky-600 disabled:bg-slate-600 disabled:cursor-not-allowed text-white font-bold py-3 px-6 rounded-md transition-all duration-300 transform hover:scale-105 disabled:scale-100"
                      >
                          {isLoading ? 'Generating...' : 'Create Presentation'}
                      </button>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-4"> {/* Added mt-4 for spacing */}
                      <div>
                          <label htmlFor="voiceover-language-select" className="block text-sm font-medium text-slate-300 mb-1">Voiceover Language</label>
                          <select
                              id="voiceover-language-select"
                              value={voiceoverLanguage}
                              onChange={(e) => setVoiceoverLanguage(e.target.value)}
                              disabled={isLoading}
                              className="w-full bg-slate-700 text-white rounded-md px-4 py-3 border border-slate-600 focus:ring-2 focus:ring-sky-500 focus:outline-none transition"
                          >
                              {voiceoverLanguages.map(lang => <option key={lang.code} value={lang.code}>{lang.name}</option>)}
                          </select>
                      </div>
                      <div>
                          <label htmlFor="voice-select" className="block text-sm font-medium text-slate-300 mb-1">Voice Style</label>
                          <select
                              id="voice-select"
                              value={voice}
                              onChange={(e) => setVoice(e.target.value)}
                              disabled={isLoading}
                              className="w-full bg-slate-700 text-white rounded-md px-4 py-3 border border-slate-600 focus:ring-2 focus:ring-sky-500 focus:outline-none transition"
                          >
                              {voices.map(v => <option key={v.id} value={v.id}>{v.name}</option>)}
                          </select>
                      </div>
                  </div>
                </div>
            </div>

            {hasResults && !isLoading && (
              <div className="mt-8 flex-1 flex flex-col animate-fade-in-up" style={{ animationFillMode: 'forwards' }}>
                <div className="bg-slate-800/50 rounded-lg p-6 shadow-2xl border border-slate-700 backdrop-blur-sm mb-6">
                    <h2 className="text-3xl font-bold text-sky-400 mb-2 text-center">{generatedPresentationTitle}</h2>
                    <p className="text-slate-300 text-center">{generatedPresentationDescription}</p>
                </div>

                <div className="flex flex-col md:flex-row justify-between items-center mb-4 gap-4">
                    <h2 className="text-2xl font-semibold text-white">Generated Presentation</h2>
                    <div className="flex items-center flex-wrap justify-center gap-3">
                        {audioUrl && (
                            <>
                                <button onClick={toggleAudio} className="flex items-center gap-2 bg-green-500 hover:bg-green-600 text-white font-semibold py-2 px-4 rounded-md transition-colors">
                                    {isPlaying ? <PauseIcon className="w-5 h-5"/> : <PlayIcon className="w-5 h-5"/>}
                                    {isPlaying ? 'Pause' : 'Play'} Voiceover
                                </button>
                                <a 
                                    href={audioUrl} 
                                    download={`${topic.replace(/\s+/g, '_')}_voiceover.wav`}
                                    className="flex items-center gap-2 bg-teal-500 hover:bg-teal-600 text-white font-semibold py-2 px-4 rounded-md transition-colors"
                                >
                                    <AudioIcon className="w-5 h-5"/>
                                    Download Audio
                                </a>
                                <audio ref={audioRef} src={audioUrl} onEnded={() => setIsPlaying(false)} onPlay={() => setIsPlaying(true)} onPause={() => setIsPlaying(false)} />
                            </>
                        )}
                        <button onClick={handleDownloadPpt} className="flex items-center gap-2 bg-indigo-500 hover:bg-indigo-600 text-white font-semibold py-2 px-4 rounded-md transition-colors">
                            <DownloadIcon className="w-5 h-5"/>
                            Download PPTX
                        </button>
                        <button 
                            onClick={handleDownloadVideo} 
                            disabled={isGeneratingVideo}
                            className="flex items-center gap-2 bg-purple-500 hover:bg-purple-600 text-white font-semibold py-2 px-4 rounded-md transition-colors disabled:bg-slate-600 disabled:cursor-not-allowed"
                        >
                            {isGeneratingVideo ? 'Generating Video...' : (
                                <>
                                    <VideoIcon className="w-5 h-5"/>
                                    Download Video
                                </>
                            )}
                        </button>
                        {videoDownloadUrl && (
                            <a 
                                href={videoDownloadUrl} 
                                download={`${topic.replace(/\s+/g, '_')}_presentation.webm`}
                                className="flex items-center gap-2 bg-pink-500 hover:bg-pink-600 text-white font-semibold py-2 px-4 rounded-md transition-colors"
                            >
                                <DownloadIcon className="w-5 h-5"/>
                                Video Ready (Download)
                            </a>
                        )}
                    </div>
                </div>

                <div className="flex-1 flex flex-col bg-slate-800/50 p-4 rounded-lg shadow-inner border border-slate-700">
                    <div className="aspect-video w-full bg-slate-900 rounded-md overflow-hidden flex justify-center items-center">
                        <SlidePreview slide={slides[currentSlide]} index={currentSlide}/>
                    </div>
                    <div className="flex justify-between items-center mt-4 text-slate-300">
                        <button 
                            onClick={() => setCurrentSlide(s => Math.max(0, s - 1))}
                            disabled={currentSlide === 0}
                            className="px-4 py-2 rounded-md bg-slate-700 hover:bg-slate-600 disabled:opacity-50 disabled:cursor-not-allowed transition"
                        >
                            Previous
                        </button>
                        <span>Slide {currentSlide + 1} of {slides.length}</span>
                        <button 
                            onClick={() => setCurrentSlide(s => Math.min(slides.length - 1, s + 1))}
                            disabled={currentSlide === slides.length - 1}
                            className="px-4 py-2 rounded-md bg-slate-700 hover:bg-slate-600 disabled:opacity-50 disabled:cursor-not-allowed transition"
                        >
                            Next
                        </button>
                    </div>
                </div>
              </div>
            )}
          </>
        )}
      </main>
    </div>
  );
};

export default App;