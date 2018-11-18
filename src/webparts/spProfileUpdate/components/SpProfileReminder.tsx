import * as React from 'react';
import { Coachmark } from 'office-ui-fabric-react/lib/Coachmark';
import { TeachingBubbleContent } from 'office-ui-fabric-react/lib/TeachingBubble';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { DirectionalHint } from 'office-ui-fabric-react/lib/common/DirectionalHint';
import { IStyle } from 'office-ui-fabric-react/lib/Styling';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { BaseComponent, classNamesFunction } from 'office-ui-fabric-react/lib/Utilities';
import { Icon } from 'office-ui-fabric-react/lib/Icon';


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
      coachmarkPosition: DirectionalHint.bottomCenter
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
            color="#EB8200"
          >
            <TeachingBubbleContent
              headline="Outstanding Profile Updates"
              hasCloseIcon={true}
              closeButtonAriaLabel="Close"
              onDismiss={this._onDismiss}
              ariaDescribedBy={'example-description1'}
              ariaLabelledBy={'example-label1'}
            >
              We need more information to target news towards you! Click the chevrons above to enter your details.
              Use the <Icon iconName="Tag"/> next to each input to get the full list of available tags.
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
  public componentWillReceiveProps(props) {
    this.setState({
      isCoachmarkVisible : props.isVisible
    });
  }

  private _onDismiss = (): void => {
    this.setState({
      isCoachmarkVisible: false
    });
  }

}