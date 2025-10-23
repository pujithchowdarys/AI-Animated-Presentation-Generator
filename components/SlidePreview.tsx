
import React from 'react';
import type { Slide } from '../types';
import { Loader } from './Loader';

interface SlidePreviewProps {
  slide: Slide;
  index: number;
}

export const SlidePreview: React.FC<SlidePreviewProps> = ({ slide, index }) => {
  return (
    <div key={index} className="w-full h-full relative flex flex-col justify-center items-center text-center bg-slate-900 animate-fade-in-up">
        {slide.image ? (
            <>
                <img src={`data:image/png;base64,${slide.image}`} alt={slide.title} className="absolute top-0 left-0 w-full h-full object-cover z-0" />
                <div className="absolute top-0 left-0 w-full h-full bg-black/60 z-10"></div>
            </>
        ) : (
          slide.imagePrompt && <Loader message="Generating image..."/>
        )}
        <div className="relative z-20 p-8 md:p-12 w-full h-full flex flex-col justify-center items-center">
            <h3 className="text-3xl md:text-4xl font-bold text-sky-400 mb-6">{slide.title}</h3>
            <ul className="text-left space-y-3 list-disc list-inside text-slate-200 text-lg md:text-xl max-w-3xl">
                {slide.content.map((point, i) => (
                <li key={i}>{point}</li>
                ))}
            </ul>
        </div>
    </div>
  );
};
