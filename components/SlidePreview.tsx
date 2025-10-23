
import React from 'react';
import type { Slide } from '../types';

interface SlidePreviewProps {
  slide: Slide;
  index: number;
}

export const SlidePreview: React.FC<SlidePreviewProps> = ({ slide, index }) => {
  return (
    <div key={index} className="w-full h-full p-8 md:p-12 flex flex-col justify-center items-center text-center bg-slate-900 animate-fade-in-up">
      <h3 className="text-3xl md:text-4xl font-bold text-sky-400 mb-6">{slide.title}</h3>
      <ul className="text-left space-y-3 list-disc list-inside text-slate-200 text-lg md:text-xl max-w-2xl">
        {slide.content.map((point, i) => (
          <li key={i}>{point}</li>
        ))}
      </ul>
    </div>
  );
};
