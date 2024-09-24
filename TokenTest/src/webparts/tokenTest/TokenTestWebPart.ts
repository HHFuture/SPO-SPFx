import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './TokenTestWebPart.module.scss';

export interface ITokenTestWebPartProps {
}

export default class TokenTestWebPart extends BaseClientSideWebPart<ITokenTestWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `<div class="${ styles.tokenTest }"></div>`;
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
