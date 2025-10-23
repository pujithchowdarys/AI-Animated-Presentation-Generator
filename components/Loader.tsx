
import React from 'react';

interface LoaderProps {
  message: string;
}

export const Loader: React.FC<LoaderProps> = ({ message }) => {
  return (
    <div className="flex flex-col items-center justify-center my-10 animate-fade-in-up">
      <div className="w-12 h-12 border-4 border-slate-500 border-t-sky-400 rounded-full animate-spinner-ease-spin mb-4"></div>
      <p className="text-slate-300 text-lg">{message}</p>
    </div>
  );
};
