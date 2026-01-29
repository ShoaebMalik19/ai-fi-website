import React, { createContext, useContext, useState } from 'react';

const AnimationContext = createContext();

export const useAnimation = () => useContext(AnimationContext);

export const AnimationProvider = ({ children }) => {
    const [isWarping, setIsWarping] = useState(false);

    const triggerWarp = (duration = 1500) => {
        setIsWarping(true);
        setTimeout(() => {
            setIsWarping(false);
        }, duration);
    };

    return (
        <AnimationContext.Provider value={{ isWarping, triggerWarp }}>
            {children}
        </AnimationContext.Provider>
    );
};
