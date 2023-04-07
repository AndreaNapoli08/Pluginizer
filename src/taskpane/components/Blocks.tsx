import * as React from 'react';
import { TextBlocks } from './Blocks/TextBlocks'
import { BulletList } from './Blocks/BulletList'
import { OrderedList} from './Blocks/OrderedList'
import { Nesting } from './Blocks/Nesting'
import { OtherBlocks } from './Blocks/OtherBlocks';

export const Blocks = () => {
    return (
        <div>
            <TextBlocks />
            <hr />
            <BulletList />
            <hr />
            <OrderedList />
            <hr />
            <Nesting />
            <hr />
            <OtherBlocks />
        </div>
    )
}