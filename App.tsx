

import React, { useState, useCallback, useRef } from 'react';
import PptxGenJS from 'pptxgenjs';
import { generatePresentationContent, generateVoiceover, generateImage } from './services/geminiService';
import type { Slide } from './types';
import { SlidePreview } from './components/SlidePreview';
import { Loader } from './components/Loader';
import { DownloadIcon, PresentationIcon, PlayIcon, PauseIcon, AudioIcon } from './components/Icons';
import { audioBufferToWavBlob, decode, decodeAudioData } from './utils/audioUtils';

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

const App: React.FC = () => {
  const [topic, setTopic] = useState<string>('');
  const [voiceoverLanguage, setVoiceoverLanguage] = useState<string>(voiceoverLanguages[0].code);
  const [voice, setVoice] = useState<string>(voices[0].id);
  const [slides, setSlides] = useState<Slide[]>([]);
  const [currentSlide, setCurrentSlide] = useState<number>(0);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [loadingMessage, setLoadingMessage] = useState<string>('');
  const [error, setError] = useState<string | null>(null);
  const [audioUrl, setAudioUrl] = useState<string | null>(null);
  const [isPlaying, setIsPlaying] = useState<boolean>(false);
  const audioRef = useRef<HTMLAudioElement>(null);

  const handleGenerate = async () => {
    if (!topic.trim()) {
      setError('Please enter a topic.');
      return;
    }
    setIsLoading(true);
    setLoadingMessage('Warming up the AI engine...');
    setError(null);
    setSlides([]);
    setAudioUrl(null);
    setCurrentSlide(0);

    try {
      setLoadingMessage('Generating presentation content...');
      const generatedSlides = await generatePresentationContent(topic, voiceoverLanguage);
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
      const audioBuffers: AudioBuffer[] = [];
      let totalLength = 0;

      for (let i = 0; i < slidesWithImages.length; i++) {
        setLoadingMessage(`Generating voice for slide ${i + 1}/${slidesWithImages.length}...`);
        const slide = slidesWithImages[i];
        if (slide.translatedSpeakerNotes) {
          const audioBase64 = await generateVoiceover(slide.translatedSpeakerNotes, voice);
          const audioBytes = decode(audioBase64);
          const audioBuffer = await decodeAudioData(audioBytes, audioContext, 24000, 1);
          audioBuffers.push(audioBuffer);
          totalLength += audioBuffer.length;
        }
      }
      
      setLoadingMessage('Stitching audio clips together...');
      const mergedBuffer = audioContext.createBuffer(1, totalLength, 24000);
      const channelData = mergedBuffer.getChannelData(0);
      let offset = 0;
      for (const buffer of audioBuffers) {
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

    slides.forEach((slide) => {
      const pptSlide = pptx.addSlide();
      
      if (slide.image) {
        pptSlide.background = { data: `data:image/png;base64,${slide.image}` };
        // Add a semi-transparent overlay for text readability
        pptSlide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '000000', transparency: 60 } });
        
        pptSlide.addText(slide.title, { 
          x: 0.5, y: 0.25, w: '90%', h: 1, 
          fontSize: 44, bold: true, color: 'FFFFFF', align: 'center',
          objectName: 'title',
          animation: { type: 'fadeIn', duration: 1, delay: 0.5 }
        });

        const contentPoints = slide.content.map(point => ({ text: point, options: { bullet: true, color: 'E2E8F0' } }));
        pptSlide.addText(contentPoints, {
          x: 1, y: 1.5, w: '80%', h: 3.5,
          fontSize: 24, color: 'FFFFFF', lineSpacing: 36,
          objectName: 'content',
          animation: { type: 'slideUp', duration: 1, delay: 0.7, effect: 'in', from: 'b' }
        });

      } else {
        pptSlide.background = { color: '1A202C' };
        pptSlide.addText(slide.title, { 
          x: 0.5, y: 0.25, w: '90%', h: 1, 
          fontSize: 36, bold: true, color: '38BDF8', align: 'center',
          objectName: 'title',
          animation: { type: 'fadeIn', duration: 1, delay: 0.5 }
        });

        const contentPoints = slide.content.map(point => ({ text: point, options: { bullet: true, color: 'E2E8F0' } }));
        pptSlide.addText(contentPoints, {
          x: 1, y: 1.5, w: '80%', h: 3.5,
          fontSize: 20, lineSpacing: 30,
          objectName: 'content',
          animation: { type: 'slideUp', duration: 1, delay: 0.7, effect: 'in', from: 'b' }
        });
      }
      pptSlide.addNotes(slide.speakerNotes);
    });

    pptx.writeFile({ fileName: `${topic.replace(/\s+/g, '_')}_presentation.pptx` });
  }, [slides, topic]);

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

  return (
    <div className="min-h-screen bg-slate-900 flex flex-col items-center p-4 sm:p-6 md:p-8">
      <header className="w-full max-w-5xl text-center mb-8">
        <div className="flex justify-center items-center gap-4 mb-2">
            <PresentationIcon className="w-10 h-10 text-sky-400" />
            <h1 className="text-3xl sm:text-4xl font-bold tracking-tight text-white">AI Animated Presentation Generator</h1>
        </div>
        <p className="text-slate-400 text-lg">Turn any topic into a downloadable presentation with voiceover in seconds.</p>
      </header>
      
      <main className="w-full max-w-5xl flex-1 flex flex-col">
        <div className="bg-slate-800/50 rounded-lg p-6 shadow-2xl border border-slate-700 backdrop-blur-sm">
            <div className="flex flex-col gap-4">
              <div className="flex flex-col sm:flex-row gap-4">
                  <input
                      type="text"
                      value={topic}
                      onChange={(e) => setTopic(e.target.value)}
                      placeholder="e.g., The Future of Renewable Energy"
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
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
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
            {error && <p className="text-red-400 mt-4 text-center">{error}</p>}
        </div>

        {isLoading && <Loader message={loadingMessage} />}
        
        {hasResults && !isLoading && (
          <div className="mt-8 flex-1 flex flex-col animate-fade-in-up" style={{ animationFillMode: 'forwards' }}>
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

      </main>
    </div>
  );
};

export default App;
