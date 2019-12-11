export interface ILandingPageState{
}

export interface ILandingPageProps {
    // These are set based on the toggles shown above the examples (not needed in real code)
    disabled?: boolean;
    checked?: boolean;
    onClick(): void;
}
