import { IBorderWebComponentProps } from "./IBorderWebComponentProps";

export interface ITitleBorderWebComponentProps extends IBorderWebComponentProps {
    title?: string;
    icon?: string;
    hideTitle?: boolean;
    titleClassName?: string;
    bodyClassName?: string;
    textColor?: string;    
    titleBkgColor?: string;
}