
export interface IModelEventProps {
    isOpen: boolean;
    title: string;
    start: string;
    end: string;
    details: string;
    color: string;
    onClose: () => void;
}
