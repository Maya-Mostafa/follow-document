export interface IfollowDocumentPropertiesProps {
    close: () => void;
    url: string;
    iframeOnLoad?: (iframe: any) => void;
    followTerm: string;
    unFollowTerm: string;
}