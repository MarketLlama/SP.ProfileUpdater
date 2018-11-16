import * as React from 'react';
import { Coachmark } from 'office-ui-fabric-react/lib/Coachmark';
import { TeachingBubbleContent } from 'office-ui-fabric-react/lib/TeachingBubble';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { DirectionalHint } from 'office-ui-fabric-react/lib/common/DirectionalHint';
import { IStyle } from 'office-ui-fabric-react/lib/Styling';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { BaseComponent, classNamesFunction } from 'office-ui-fabric-react/lib/Utilities';

export interface ISpProfileReminderState {
  isCoachmarkVisible?: boolean;
  coachmarkPosition: DirectionalHint;
}

export interface ISpProfileReminderProps {
    target : HTMLDivElement;
    isVisible : boolean;
}

export interface ISpProfileReminderStyles {
  /**
   * Style for the root element in the default enabled/unchecked state.
   */
  root?: IStyle;

  /**
   * The example button container
   */
  buttonContainer: IStyle;

  /**
   * The dropdown component container
   */
  dropdownContainer: IStyle;
}

export class SpProfileReminder extends React.Component<ISpProfileReminderProps, ISpProfileReminderState> {

  public constructor(props: ISpProfileReminderProps) {
    super(props);

    this.state = {
      isCoachmarkVisible: false,
      coachmarkPosition: DirectionalHint.bottomAutoEdge
    };
  }

  public render(): JSX.Element {
    const { isCoachmarkVisible } = this.state;

    const getClassNames = classNamesFunction<{}, ISpProfileReminderStyles>();
    const classNames = getClassNames(() => {
      return {
        buttonContainer: {
          marginTop: '30px',
          display: 'inline-block'
        }
      };
    }, {});

    const buttonProps: IButtonProps = {
      text: 'Try it'
    };

    const buttonProps2: IButtonProps = {
      text: 'Try it again'
    };

    return (
      <div className={classNames.root}>
        {isCoachmarkVisible && (
          <Coachmark
            target={this.props.target}
            positioningContainerProps={{
              directionalHint: this.state.coachmarkPosition
            }}
            ariaAlertText="A Coachmark has appeared"
            ariaDescribedBy={'coachmark-desc1'}
            ariaLabelledBy={'coachmark-label1'}
            ariaDescribedByText={'Press enter or alt + C to open the Coachmark notification'}
            ariaLabelledByText={'Coachmark notification'}
          >
            <TeachingBubbleContent
              headline="Example Title"
              hasCloseIcon={true}
              closeButtonAriaLabel="Close"
              primaryButtonProps={buttonProps}
              secondaryButtonProps={buttonProps2}
              onDismiss={this._onDismiss}
              ariaDescribedBy={'example-description1'}
              ariaLabelledBy={'example-label1'}
            >
              Welcome to the land of Coachmarks!
            </TeachingBubbleContent>
          </Coachmark>
        )}
      </div>
    );
  }

  public componentDidMount(){
    this.setState({
        isCoachmarkVisible: this.props.isVisible
    });
  }

  private _onDismiss = (): void => {
    this.setState({
      isCoachmarkVisible: false
    });
  }

}