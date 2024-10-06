declare module 'win-process-audio-capture' {
    /**
     * Starts the audio capture for a specific process.
     * @param processId The ID of the process to capture audio from.
     * @returns A string representation of the HRESULT.
     */
    function startCapture(processId: number): string;

    /**
     * Gets the captured audio data.
     * @returns A string containing debug information about the capture state.
     */
    function getAudioData(): string;

    /**
     * Starts the default audio capture.
     * @returns void
     */
    function defaultCaputure(): void;

    /**
     * Gets the captured audio data.
     * @returns A Buffer containing the audio data, or null if no data is available.
     */
    function getpcapture(): Buffer | null;

    /**
     * Gets the number of frames captured.
     * @returns The number of frames captured.
     */
    function getnNumFramesCaptured(): number;

    /**
     * Gets the flags associated with the captured audio.
     * @returns A string containing information about the flags.
     */
    function getdwFlags(): string;

    /**
     * Gets the number of available frames.
     * @returns The number of available frames.
     */
    function getnNumFramesAvailable(): number;

    /**
     * Gets the audio format information.
     * @returns An object containing audio format details.
     */
    function getAudioFormat(): {
        channels: number;
        sampleRate: number;
        bitsPerSample: number;
        avgBytesPerSec: number;
    };
}