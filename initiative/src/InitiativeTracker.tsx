import {useInitiativeSize} from './useInitiativeSize';
import {createLogEntry,publishLogEntry} from '../logTransport.js';
import { useEffect, useMemo, useRef, useState } from "react";

import Stack from "@mui/material/Stack";
import IconButton from "@mui/material/IconButton";
import List from "@mui/material/List";
import Box from "@mui/material/Box";
import Tooltip from "@mui/material/Tooltip";

import SkipNextRounded from "@mui/icons-material/SkipNextRounded";
import PlaylistAddRounded from "@mui/icons-material/PlaylistAddRounded";
import CasinoRounded from "@mui/icons-material/CasinoRounded";

import OBR, { Image, isImage, Item, Player } from "@owlbear-rodeo/sdk";

import { InitiativeItem } from "./InitiativeItem";




import { InitiativeListItem } from "./InitiativeListItem";
import { getPluginId } from "./getPluginId";
import { InitiativeHeader } from "./InitiativeHeader";
import { isPlainObject } from "./isPlainObject";
import { focusInitiativeItem } from "./focusInitiativeItem";

import {getOdysseyData} from '../../shared.js';

const ODYSSEY_META_KEY = "com.codex.body-hp/data";

function parseInitiativeCount(count: string) {
  const numeric = Number.parseFloat(count);
  return Number.isFinite(numeric) ? numeric : Number.NEGATIVE_INFINITY;
}

function compareInitiative(left: InitiativeItem, right: InitiativeItem) {
  const difference = parseInitiativeCount(right.count) - parseInitiativeCount(left.count);
  if (difference !== 0) return difference;
  return left.name.localeCompare(right.name);
}

function isTrackableToken(item: Item): item is Image {
  return isImage(item) && (item.layer === "CHARACTER" || item.layer === "MOUNT");
}

function getReaction(item: Item,items:Item[]) {return getOdysseyData(item,items).attributes.Reaction||0;}

/** Check that the item metadata is in the correct format */
function isMetadata(
  metadata: unknown
): metadata is { count: string; active: boolean } {
  return (
    isPlainObject(metadata) &&
    typeof metadata.count === "string" &&
    typeof metadata.active === "boolean"
  );
}

export function InitiativeTracker() {
  const {size,handle:resizeHandle}=useInitiativeSize();
  const [initiativeItems, setInitiativeItems] = useState<InitiativeItem[]>([]);
  const [role, setRole] = useState<"GM" | "PLAYER">("PLAYER");
  const [status, setStatus] = useState("Select character tokens and add them to initiative.");

  useEffect(() => {
    const handlePlayerChange = (player: Player) => {
      setRole(player.role);
    };
    OBR.player.getRole().then(setRole);
    return OBR.player.onChange(handlePlayerChange);
  }, []);

  useEffect(() => {
    const handleItemsChange = async (items: Item[]) => {
      const initiativeItems: InitiativeItem[] = [];
      for (const item of items) {
        if (isTrackableToken(item)) {
          const metadata = item.metadata[getPluginId("metadata")];
          if (isMetadata(metadata)) {
            initiativeItems.push({
              id: item.id,
              imageUrl:item.image.url,
              count: metadata.count,
              name: item.text?.plainText || item.name,
              active: metadata.active,
              visible: item.visible,
            });
          }
        }
      }
      setInitiativeItems(initiativeItems);
    };

    OBR.scene.items.getItems().then(handleItemsChange);
    return OBR.scene.items.onChange(handleItemsChange);
  }, []);

  const sortedInitiativeItems = useMemo(
    () => [...initiativeItems].sort(compareInitiative),
    [initiativeItems]
  );

  async function handleRemoveClick(id:string){
    try{if(role!=='GM'||await OBR.player.getRole()!=='GM')throw Error('Only the GM can remove participants.');
      await OBR.scene.items.updateItems([id],items=>{for(const item of items)delete item.metadata[getPluginId('metadata')];});
      setStatus('Removed from initiative.');
    }catch(error){setStatus('Unable to remove participant.');await OBR.notification.show((error as Error).message||'Unable to remove participant.','ERROR');}
  }
  async function handleNextClick() {
    if (role !== "GM" || sortedInitiativeItems.length === 0) return;

    const activeIndex = sortedInitiativeItems.findIndex((initiative) => initiative.active);
    const nextIndex = activeIndex >= 0 ? (activeIndex + 1) % sortedInitiativeItems.length : 0;
    const nextId = sortedInitiativeItems[nextIndex]?.id;
    if (!nextId) return;

    setInitiativeItems((prev) =>
      prev.map((initiative) => ({
        ...initiative,
        active: initiative.id === nextId,
      }))
    );

    await OBR.scene.items.updateItems(
      sortedInitiativeItems.map((initiative) => initiative.id),
      (items) => {
        for (const item of items) {
          const metadata = item.metadata[getPluginId("metadata")];
          if (isMetadata(metadata)) {
            metadata.active = item.id === nextId;
          }
        }
      }
    );

    await focusInitiativeItem(nextId);
    setStatus("Moved to the next token in initiative.");
  }

  function handleInitiativeCountChange(id: string, newCount: string) {
    if (role !== "GM") return;

    // Set local items immediately
    setInitiativeItems((prev) =>
      prev.map((initiative) => {
        if (initiative.id === id) {
          return {
            ...initiative,
            count: newCount,
          };
        } else {
          return initiative;
        }
      })
    );
    // Sync changes over the network
    OBR.scene.items.updateItems([id], (items) => {
      for (let item of items) {
        const metadata = item.metadata[getPluginId("metadata")];
        if (isMetadata(metadata)) {
          metadata.count = newCount;
        }
      }
    });
  }

  async function handleAddSelectedClick() {
    if (role !== "GM") return;

    const [selection, items] = await Promise.all([
      OBR.player.getSelection(),
      OBR.scene.items.getItems(),
    ]);
    const selectedIds = new Set(selection ?? []);
    const selectedTokens = items.filter(
      (item) => selectedIds.has(item.id) && isTrackableToken(item)
    );

    if (selectedTokens.length === 0) {
      setStatus("Select one or more character tokens first.");
      return;
    }

    await OBR.scene.items.updateItems(
      selectedTokens.map((item) => item.id),
      (items) => {
        for (const item of items) {
          if (item.metadata[getPluginId("metadata")] === undefined) {
            item.metadata[getPluginId("metadata")] = {
              count: "0",
              active: false,
            };
          }
        }
      }
    );

    setStatus(`Added ${selectedTokens.length} token${selectedTokens.length === 1 ? "" : "s"} to initiative.`);
  }

  async function handleRollInitiativeClick() {
    if (role !== "GM" || initiativeItems.length === 0) return;

    const items = await OBR.scene.items.getItems();
    const trackedIds = new Set(initiativeItems.map((initiative) => initiative.id));
    const rolls = new Map<string, string>();
    const logLines: string[] = [];

    for (const item of items) {
      if (!trackedIds.has(item.id) || !isTrackableToken(item)) continue;
      const die = Math.floor(Math.random() * 20) + 1, reaction = getReaction(item,items);
      const initiativeRoll = die + reaction;
      logLines.push(`${item.name || item.id}: d20 (${die}) + Reaction ${reaction} = ${initiativeRoll}`);
      rolls.set(item.id, String(initiativeRoll));
    }

    if (rolls.size === 0) {
      setStatus("No valid tracked character tokens found for initiative.");
      return;
    }

    const nextActiveId =
      [...initiativeItems]
        .map((initiative) => ({
          ...initiative,
          count: rolls.get(initiative.id) ?? initiative.count,
        }))
        .sort(compareInitiative)[0]?.id ?? null;

    setInitiativeItems((prev) =>
      prev.map((initiative) => ({
        ...initiative,
        count: rolls.get(initiative.id) ?? initiative.count,
        active: initiative.id === nextActiveId,
      }))
    );

    await OBR.scene.items.updateItems(Array.from(trackedIds), (items) => {
      for (const item of items) {
        const metadata = item.metadata[getPluginId("metadata")];
        if (isMetadata(metadata)) {
          metadata.count = rolls.get(item.id) ?? metadata.count;
          metadata.active = item.id === nextActiveId;
        }
      }
    });

    try {await publishLogEntry(OBR,createLogEntry('Initiative · d20 + Reaction',logLines.join('\n'),'success'),true);}
    catch {setStatus('Initiative saved; public log unavailable.');return;}
    if (nextActiveId) {
      await focusInitiativeItem(nextActiveId);
    }
    setStatus("Rolled initiative for all tracked tokens using d20 + Reaction.");
  }

  const listRef=useRef<HTMLUListElement>(null);
  const activeId=initiativeItems.find(item=>item.active)?.id;
  const activePosition=sortedInitiativeItems.findIndex(item=>item.id===activeId);

  useEffect(()=>{listRef.current?.querySelector<HTMLElement>('[data-active="true"]')?.scrollIntoView({block:'nearest',inline:'nearest',behavior:'smooth'});},[activeId,activePosition,role]);

  return (
    <Stack direction="row" height="100vh" sx={{overflow:"hidden",minWidth:0,alignItems:"center",'--initiative-avatar-size':Math.max(36,Math.min(120,size.height-64))+'px'}}>
      <InitiativeHeader
        action={
          <Stack direction="column" spacing={0}>
            <Tooltip disableInteractive title={role === "GM" ? "Add Selected Tokens" : "GM only"}>
              <span>
                <IconButton
                  size="small"
                  aria-label="add-selected"
                  onClick={handleAddSelectedClick}
                  disabled={role !== "GM"}
                >
                  <PlaylistAddRounded />
                </IconButton>
              </span>
            </Tooltip>
            <Tooltip disableInteractive title={role === "GM" ? "Roll Initiative (d20 + Reaction)" : "GM only"}>
              <span>
                <IconButton
                  size="small"
                  aria-label="roll-initiative"
                  onClick={handleRollInitiativeClick}
                  disabled={role !== "GM" || initiativeItems.length === 0}
                >
                  <CasinoRounded />
                </IconButton>
              </span>
            </Tooltip>
            <Tooltip disableInteractive title={role === "GM" ? "Next Token" : "GM only"}>
              <span>
                <IconButton
                  size="small"
                  aria-label="next"
                  onClick={handleNextClick}
                  disabled={role !== "GM" || initiativeItems.length === 0}
                >
                  <SkipNextRounded />
                </IconButton>
              </span>
            </Tooltip>
          </Stack>
        }
      />
      <Box sx={{minHeight:0,minWidth:0,flex:1,overflow:"hidden"}}>
        <List ref={listRef} onWheel={event=>{if(Math.abs(event.deltaY)>Math.abs(event.deltaX))event.currentTarget.scrollLeft+=event.deltaY;}} className="initiative-strip" aria-label="Initiative order" sx={{display:"flex",flexWrap:"nowrap",gap:.75,p:.5,alignItems:"center",overflowX:"auto",overflowY:"hidden",minHeight:88}}>
          {sortedInitiativeItems.map((initiative) => (
              <InitiativeListItem
                key={initiative.id}
                initiative={initiative}
                onCountChange={(newCount) => {
                  handleInitiativeCountChange(initiative.id, newCount);
                }}
                onRemove={()=>void handleRemoveClick(initiative.id)}
                showHidden={role === "GM"}
                editable={role === "GM"}
              />
            ))}
        </List>
      </Box>
      <span className="initiative-sr-only" role="status">{status}</span>{resizeHandle}
    </Stack>
  );
}
