import { IBorderProps } from "./IBorderProps";

export interface ITitleBorderProps extends IBorderProps {
    title?: string;
    icon?: string;
    hideTitle?: boolean;
    titleClassName?: string;
    bodyClassName?: string;
    textColor?: string;
    titleBkgColor?: string;
}