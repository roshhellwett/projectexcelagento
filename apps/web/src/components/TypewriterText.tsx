import React, { useState, useEffect } from 'react';

interface TypewriterTextProps {
  text: string;
  animate?: boolean;
  speed?: number; // ms per chunk
  onComplete?: () => void;
}

export const TypewriterText: React.FC<TypewriterTextProps> = ({
  text,
  animate = true,
  speed = 12,
  onComplete,
}) => {
  const [displayedLength, setDisplayedLength] = useState(animate ? 0 : text.length);
  const [isFinished, setIsFinished] = useState(!animate);

  useEffect(() => {
    if (!animate) {
      setDisplayedLength(text.length);
      setIsFinished(true);
      return;
    }

    setDisplayedLength(0);
    setIsFinished(false);

    let currentIndex = 0;
    const interval = setInterval(() => {
      currentIndex += 2; // Type 2 chars per tick for smooth fast typing
      if (currentIndex >= text.length) {
        setDisplayedLength(text.length);
        setIsFinished(true);
        clearInterval(interval);
        onComplete?.();
      } else {
        setDisplayedLength(currentIndex);
      }
    }, speed);

    return () => clearInterval(interval);
  }, [text, animate, speed, onComplete]);

  const displayedText = text.slice(0, displayedLength);

  // Render text with basic bolding and line breaks
  const renderFormatted = (str: string) => {
    const parts = str.split('\n');
    return parts.map((line, lIdx) => {
      // Bold **text**
      const boldParts = line.split(/(\*\*.*?\*\*)/g);
      return (
        <React.Fragment key={lIdx}>
          {boldParts.map((part, pIdx) => {
            if (part.startsWith('**') && part.endsWith('**')) {
              return <strong key={pIdx}>{part.slice(2, -2)}</strong>;
            }
            return part;
          })}
          {lIdx < parts.length - 1 && <br />}
        </React.Fragment>
      );
    });
  };

  return (
    <div
      className="typewriter-content"
      onClick={() => {
        // Skip animation on click
        if (!isFinished) {
          setDisplayedLength(text.length);
          setIsFinished(true);
        }
      }}
    >
      {renderFormatted(displayedText)}
      {!isFinished && <span className="typing-cursor">▋</span>}
    </div>
  );
};
